#!/usr/bin/env python3
"""template_store.py -- テンプレート（ベースpptx + テーマJSON）の保管層.

millet_ms 側の永続化（CloudSQL / Firestore）に差し替えられるよう、
インターフェースを切って実装を隠蔽する。millet_ms では TemplateStore を
実装したクラスを注入すればよく、このモジュールを import する必要はない。

レコードの構造:
  tenant_id   会社ID（マルチテナントの分離単位）
  template_id テンプレートID
  version     1から。**発行済みのバージョンは書き換えない**。
              過去に生成した資料を再現できるようにするため
  status      draft | published
  theme       デザイントークン（dict）
  base        ベースpptx（マスター/レイアウトのみ。顧客のスライドは含まない）

ベースpptxは python-pptx がローカルパスを要求するため、取得時に
プロセス内キャッシュへ展開する。バッチで同じテンプレートを何度も使うので
ダウンロードは1回で済ませる。
"""
import abc
import hashlib
import json
import pathlib
import shutil
import tempfile

CACHE_DIR = pathlib.Path(tempfile.gettempdir()) / "md2pptx-templates"


class TemplateNotFound(Exception):
    pass


class TemplateStore(abc.ABC):
    @abc.abstractmethod
    def save(self, tenant_id, template_id, version, theme, base_bytes, status="draft"):
        """新しいバージョンを書き込む -> レコード(dict)"""

    @abc.abstractmethod
    def load(self, tenant_id, template_id, version=None):
        """version省略時は最新のpublished（無ければ最新のdraft） -> レコード(dict)"""

    @abc.abstractmethod
    def list(self, tenant_id):
        """テンプレートの一覧（最新バージョンのみ） -> [レコード]"""

    @abc.abstractmethod
    def publish(self, tenant_id, template_id, version):
        """指定バージョンを published にする"""

    # ---- 共通処理 ----

    def materialize(self, record):
        """ベースpptxをローカルへ展開し、theme.base_pptx をそのパスに差し替える.

        同一内容なら再利用する（バッチで同じテンプレートを何度も開くため）。
        """
        data = record["base"]
        digest = hashlib.sha256(data).hexdigest()[:16]
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        path = CACHE_DIR / f"{digest}.pptx"
        if not path.is_file():
            tmp = path.with_suffix(".part")
            tmp.write_bytes(data)
            tmp.replace(path)          # 途中で落ちても壊れたファイルを残さない
        theme = dict(record["theme"])
        theme["base_pptx"] = str(path)
        return theme

    def next_version(self, tenant_id, template_id):
        try:
            cur = self.load(tenant_id, template_id, version="latest")
            return cur["version"] + 1
        except TemplateNotFound:
            return 1


class LocalTemplateStore(TemplateStore):
    """ファイルシステム実装。開発とテスト用."""

    def __init__(self, root):
        self.root = pathlib.Path(root)

    def _dir(self, tenant_id, template_id, version=None):
        d = self.root / tenant_id / template_id
        return d / f"v{version}" if version else d

    def save(self, tenant_id, template_id, version, theme, base_bytes, status="draft"):
        d = self._dir(tenant_id, template_id, version)
        d.mkdir(parents=True, exist_ok=True)
        (d / "base.pptx").write_bytes(base_bytes)
        meta = {"tenant_id": tenant_id, "template_id": template_id,
                "version": version, "status": status, "theme": theme}
        (d / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1),
                                     encoding="utf-8")
        return {**meta, "base": base_bytes}

    def _versions(self, tenant_id, template_id):
        d = self._dir(tenant_id, template_id)
        if not d.is_dir():
            return []
        return sorted((int(p.name[1:]) for p in d.glob("v*") if p.is_dir()))

    def load(self, tenant_id, template_id, version=None):
        vs = self._versions(tenant_id, template_id)
        if not vs:
            raise TemplateNotFound(f"{tenant_id}/{template_id}")
        if version in (None, "published", "latest"):
            metas = [json.loads((self._dir(tenant_id, template_id, v) / "meta.json")
                                .read_text(encoding="utf-8")) for v in vs]
            if version == "latest":
                chosen = metas[-1]
            else:
                pub = [m for m in metas if m["status"] == "published"]
                chosen = pub[-1] if pub else metas[-1]
        else:
            version = int(version)
            if version not in vs:
                raise TemplateNotFound(f"{tenant_id}/{template_id} v{version}")
            chosen = json.loads((self._dir(tenant_id, template_id, version) / "meta.json")
                                .read_text(encoding="utf-8"))
        base = (self._dir(tenant_id, template_id, chosen["version"]) / "base.pptx")
        return {**chosen, "base": base.read_bytes()}

    def list(self, tenant_id):
        d = self.root / tenant_id
        if not d.is_dir():
            return []
        out = []
        for t in sorted(p.name for p in d.iterdir() if p.is_dir()):
            rec = self.load(tenant_id, t, version="latest")
            rec.pop("base")            # 一覧では中身を返さない
            out.append(rec)
        return out

    def publish(self, tenant_id, template_id, version):
        path = self._dir(tenant_id, template_id, int(version)) / "meta.json"
        if not path.is_file():
            raise TemplateNotFound(f"{tenant_id}/{template_id} v{version}")
        meta = json.loads(path.read_text(encoding="utf-8"))
        meta["status"] = "published"
        path.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding="utf-8")
        return meta


class GcsTemplateStore(TemplateStore):
    """GCS実装。メタはJSONオブジェクトとして隣に置く.

    millet_ms に組み込む際は、メタだけをCloudSQL/Firestoreに移して
    一覧・検索・監査に使うのが素直（このクラスを差し替える）。
    """

    def __init__(self, bucket, prefix="templates"):
        self.bucket_name, self.prefix = bucket, prefix.strip("/")
        self._client = None

    @property
    def bucket(self):
        if self._client is None:
            from google.cloud import storage
            self._client = storage.Client()
        return self._client.bucket(self.bucket_name)

    def _key(self, tenant_id, template_id, version, name):
        return f"{self.prefix}/{tenant_id}/{template_id}/v{version}/{name}"

    def save(self, tenant_id, template_id, version, theme, base_bytes, status="draft"):
        meta = {"tenant_id": tenant_id, "template_id": template_id,
                "version": version, "status": status, "theme": theme}
        self.bucket.blob(self._key(tenant_id, template_id, version, "base.pptx")) \
            .upload_from_string(base_bytes)
        self.bucket.blob(self._key(tenant_id, template_id, version, "meta.json")) \
            .upload_from_string(json.dumps(meta, ensure_ascii=False),
                               content_type="application/json")
        return {**meta, "base": base_bytes}

    def _metas(self, tenant_id, template_id):
        pre = f"{self.prefix}/{tenant_id}/{template_id}/"
        out = []
        for b in self.bucket.list_blobs(prefix=pre):
            if b.name.endswith("meta.json"):
                out.append(json.loads(b.download_as_bytes()))
        return sorted(out, key=lambda m: m["version"])

    def load(self, tenant_id, template_id, version=None):
        metas = self._metas(tenant_id, template_id)
        if not metas:
            raise TemplateNotFound(f"{tenant_id}/{template_id}")
        if version in (None, "published"):
            pub = [m for m in metas if m["status"] == "published"]
            chosen = pub[-1] if pub else metas[-1]
        elif version == "latest":
            chosen = metas[-1]
        else:
            match = [m for m in metas if m["version"] == int(version)]
            if not match:
                raise TemplateNotFound(f"{tenant_id}/{template_id} v{version}")
            chosen = match[0]
        base = self.bucket.blob(
            self._key(tenant_id, template_id, chosen["version"], "base.pptx")
        ).download_as_bytes()
        return {**chosen, "base": base}

    def list(self, tenant_id):
        pre = f"{self.prefix}/{tenant_id}/"
        ids = {b.name[len(pre):].split("/")[0]
               for b in self.bucket.list_blobs(prefix=pre)}
        out = []
        for t in sorted(ids):
            rec = self.load(tenant_id, t, version="latest")
            rec.pop("base")
            out.append(rec)
        return out

    def publish(self, tenant_id, template_id, version):
        blob = self.bucket.blob(
            self._key(tenant_id, template_id, int(version), "meta.json"))
        if not blob.exists():
            raise TemplateNotFound(f"{tenant_id}/{template_id} v{version}")
        meta = json.loads(blob.download_as_bytes())
        meta["status"] = "published"
        blob.upload_from_string(json.dumps(meta, ensure_ascii=False),
                                content_type="application/json")
        return meta


def from_env(env):
    """環境変数から実装を選ぶ。TEMPLATE_BUCKET があればGCS、無ければローカル."""
    bucket = env.get("TEMPLATE_BUCKET")
    if bucket:
        return GcsTemplateStore(bucket, env.get("TEMPLATE_PREFIX", "templates"))
    return LocalTemplateStore(env.get("TEMPLATE_ROOT", "./.templates"))


def clear_cache():
    shutil.rmtree(CACHE_DIR, ignore_errors=True)
