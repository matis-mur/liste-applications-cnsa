"""
Export Confluence -> Excel.

Ce fichier contient tout le moteur d'export ET le point d'entrée HTTP.
- Sur Vercel, il est servi automatiquement à l'adresse /api/export.
- En local, serveur_local.py réutilise exactement le même code.

Le client envoie (POST JSON) : url, email, token, filename (optionnel).
La réponse est le fichier .xlsx, ou un JSON {"error": "..."} en cas de problème.
Les identifiants ne sont jamais écrits sur disque ni dans les logs.
"""

from __future__ import annotations

import io
import json
import os
import re
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, quote, unquote_plus, urljoin, urlparse

import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:  # fuseau horaire pour afficher les dates en heure de Paris
    from zoneinfo import ZoneInfo

    PARIS = ZoneInfo("Europe/Paris")
except Exception:  # pragma: no cover - environnement sans base de fuseaux
    PARIS = timezone.utc


# --------------------------------------------------------------------------- #
# Réglages
# --------------------------------------------------------------------------- #

ON_VERCEL = bool(os.environ.get("VERCEL"))
# Temps maximum consacré à l'export (Vercel coupe la fonction à 60 s).
MAX_SECONDS = float(os.environ.get("EXPORT_MAX_SECONDS", "52" if ON_VERCEL else "600"))
# Vercel refuse les réponses de plus de 4,5 Mo.
MAX_RESPONSE_BYTES = int(os.environ.get("EXPORT_MAX_BYTES", str(4_300_000 if ON_VERCEL else 200_000_000)))
MAX_ITEMS = int(os.environ.get("EXPORT_MAX_ITEMS", "3000"))
MAX_TABLE_SHEETS = 150
EXCEL_CELL_LIMIT = 32_000

# Par sécurité, le serveur n'appelle que des sites Confluence Cloud (*.atlassian.net).
# EXPORT_ALLOWED_HOSTS permet d'ajouter des domaines personnalisés (séparés par des virgules).
EXTRA_HOSTS = [h.strip().lower() for h in os.environ.get("EXPORT_ALLOWED_HOSTS", "").split(",") if h.strip()]
ALLOW_INSECURE = os.environ.get("EXPORT_ALLOW_INSECURE_HOSTS") == "1"  # tests uniquement

# Configuration « bouton unique » : si l'email et le token sont définis côté serveur
# (variables d'environnement Vercel), les visiteurs n'ont rien à saisir.
CONF_EMAIL = os.environ.get("CONFLUENCE_EMAIL", "").strip()
CONF_TOKEN = os.environ.get("CONFLUENCE_TOKEN", "").strip()
SOURCE_URL = os.environ.get(
    "CONFLUENCE_URL",
    "https://cnsa.atlassian.net/wiki/spaces/Urbanisati/pages/1316880408/06.+Fiche+applicative",
).strip()
ACCESS_CODE = os.environ.get("ACCESS_CODE", "").strip()
DEFAULT_FILENAME = os.environ.get("EXPORT_FILENAME", "Liste des Applications CNSA").strip()
EXPORT_FORMAT = os.environ.get("EXPORT_FORMAT", "fiches").strip().lower()  # fiches | pages

GREEN = "7AB51D"
GREEN_DARK = "5E9412"
GREEN_SOFT = "EEF6E0"


# --------------------------------------------------------------------------- #
# Erreurs lisibles par l'utilisateur
# --------------------------------------------------------------------------- #


class ExportError(Exception):
    def __init__(self, message: str, status: int = 400):
        super().__init__(message)
        self.message = message
        self.status = status


# --------------------------------------------------------------------------- #
# Lecture de l'URL saisie
# --------------------------------------------------------------------------- #


@dataclass
class Target:
    origin: str  # https://xxx.atlassian.net
    context: str  # "/wiki"
    kind: str  # page | folder | space | tiny | title
    id: str | None = None
    space_key: str | None = None
    title: str | None = None
    original_url: str = ""

    @property
    def wiki(self) -> str:
        return self.origin + self.context


_CONTAINER_TYPES = {
    "pages": "page",
    "page": "page",
    "folder": "folder",
    "folders": "folder",
    "whiteboard": "whiteboard",
    "whiteboards": "whiteboard",
    "database": "database",
    "databases": "database",
    "embed": "embed",
    "embeds": "embed",
}


def parse_confluence_url(raw: str) -> Target:
    url = (raw or "").strip()
    if not url:
        raise ExportError("Collez l'URL du dossier Confluence.")
    if not re.match(r"^https?://", url, re.I):
        url = "https://" + url
    parsed = urlparse(url)
    host = (parsed.hostname or "").lower()
    if not host:
        raise ExportError("Cette URL n'est pas valide.")

    allowed = host.endswith(".atlassian.net") or host in EXTRA_HOSTS
    if ALLOW_INSECURE and host in ("localhost", "127.0.0.1"):
        allowed = True
    if not allowed:
        raise ExportError(
            "Seuls les espaces Confluence Cloud (adresse en xxx.atlassian.net) sont acceptés. "
            "Vérifiez l'URL copiée."
        )
    scheme = parsed.scheme.lower()
    if scheme != "https" and not (ALLOW_INSECURE and host in ("localhost", "127.0.0.1")):
        scheme = "https"
    origin = f"{scheme}://{parsed.netloc}"

    parts = [p for p in parsed.path.split("/") if p]
    context = ""
    if parts and parts[0].lower() == "wiki":
        context = "/wiki"
        parts = parts[1:]
    elif host.endswith(".atlassian.net"):
        context = "/wiki"

    def target(kind, **kw):
        return Target(origin=origin, context=context, kind=kind, original_url=url, **kw)

    query = parse_qs(parsed.query)
    if "pageId" in query and query["pageId"][0].isdigit():
        return target("page", id=query["pageId"][0])

    # https://xxx.atlassian.net/wiki/x/AbCd  (lien court)
    if len(parts) >= 2 and parts[0] == "x":
        return target("tiny", id=parts[1])

    if parts and parts[0] == "spaces" and len(parts) >= 2:
        key = parts[1]
        rest = parts[2:]
        # /spaces/KEY/pages/edit-v2/123  ou  /spaces/KEY/pages/123/Titre
        if rest and rest[0] in _CONTAINER_TYPES:
            ids = [p for p in rest[1:] if p.isdigit()]
            if ids:
                kind = _CONTAINER_TYPES[rest[0]]
                if kind not in ("page", "folder"):
                    raise ExportError(
                        "Ce lien pointe vers un tableau blanc ou une base de données. "
                        "Collez plutôt le lien d'une page ou d'un dossier."
                    )
                return target(kind, id=ids[0], space_key=key)
        # /spaces/KEY  ou  /spaces/KEY/overview : tout l'espace
        if not rest or rest[0] in ("overview", "pages"):
            return target("space", space_key=key)
        raise ExportError("Ce lien Confluence n'est pas reconnu. Copiez l'URL d'un dossier ou d'une page.")

    # Ancien format : /display/KEY/Titre+de+la+page
    if parts and parts[0] == "display" and len(parts) >= 2:
        if len(parts) == 2:
            return target("space", space_key=parts[1])
        return target("title", space_key=parts[1], title=unquote_plus(parts[2]))

    raise ExportError("Ce lien Confluence n'est pas reconnu. Copiez l'URL d'un dossier ou d'une page.")


# --------------------------------------------------------------------------- #
# Client HTTP Confluence
# --------------------------------------------------------------------------- #


class ConfluenceClient:
    def __init__(self, target: Target, email: str, token: str, deadline: float):
        self.t = target
        self.deadline = deadline
        self._last_link_next = None
        self.session = requests.Session()
        self.session.auth = (email, token)
        self.session.headers.update(
            {"Accept": "application/json", "User-Agent": "export-confluence-excel/1.0"}
        )

    def api(self, path: str) -> str:
        return f"{self.t.wiki}/api/v2{path}"

    def _check_time(self):
        if time.monotonic() > self.deadline:
            raise ExportError(
                "Le dossier est trop volumineux pour être exporté en une fois. "
                "Essayez avec un sous-dossier plus petit.",
                413,
            )

    def request(self, method: str, url: str, *, params=None, json_body=None, allow_404=False, allow_redirects=True):
        for attempt in range(4):
            self._check_time()
            remaining = max(5.0, self.deadline - time.monotonic())
            try:
                r = self.session.request(
                    method,
                    url,
                    params=params,
                    json=json_body,
                    timeout=(10, min(45.0, remaining)),
                    allow_redirects=allow_redirects,
                )
            except requests.exceptions.SSLError:
                raise ExportError("Connexion sécurisée impossible avec ce site Confluence.", 502)
            except requests.exceptions.ConnectionError:
                raise ExportError(
                    "Impossible de joindre ce site Confluence. Vérifiez l'adresse dans l'URL.", 502
                )
            except requests.exceptions.Timeout:
                if attempt < 2:
                    continue
                raise ExportError("Confluence met trop de temps à répondre. Réessayez dans un instant.", 504)

            if r.status_code == 429 or r.status_code >= 500:
                if attempt < 3:
                    wait = r.headers.get("Retry-After")
                    try:
                        delay = min(float(wait), 8.0) if wait else 1.5 * (attempt + 1)
                    except ValueError:
                        delay = 2.0
                    time.sleep(delay)
                    continue
            return self._handle(r, allow_404)
        return self._handle(r, allow_404)

    @staticmethod
    def _handle(r: requests.Response, allow_404: bool):
        if r.status_code == 401:
            raise ExportError("Connexion refusée : vérifiez l'email et le token API.", 401)
        if r.status_code == 403:
            raise ExportError("Ce compte n'a pas le droit de consulter ce contenu.", 403)
        if r.status_code == 404:
            if allow_404:
                return None
            raise ExportError(
                "Contenu introuvable : vérifiez l'URL, ou que votre compte a bien accès à ce dossier.", 404
            )
        if r.status_code == 429:
            raise ExportError("Confluence limite le nombre de requêtes. Réessayez dans une minute.", 429)
        if r.status_code >= 400:
            raise ExportError(f"Confluence a renvoyé une erreur ({r.status_code}). Réessayez plus tard.", 502)
        return r

    def get_json(self, path_or_url: str, params=None, allow_404=False):
        url = path_or_url if path_or_url.startswith("http") else self.api(path_or_url)
        r = self.request("GET", url, params=params, allow_404=allow_404)
        self._last_link_next = None
        if r is None:
            return None
        m = re.search(r'<([^>]+)>;\s*rel="?next"?', r.headers.get("Link", ""))
        if m:
            self._last_link_next = m.group(1)
        try:
            return r.json()
        except ValueError:
            raise ExportError(
                "La réponse de Confluence est illisible. Vérifiez qu'il s'agit bien d'un Confluence Cloud.", 502
            )

    def paginate(self, path: str, params=None, allow_404=False):
        url = self.api(path)
        first = True
        while url:
            data = self.get_json(url, params=params if first else None, allow_404=allow_404 and first)
            if data is None:
                return
            first = False
            for item in data.get("results", []) or []:
                yield item
            nxt = (data.get("_links") or {}).get("next") or self._last_link_next
            url = urljoin(self.t.origin + "/", nxt) if nxt else None


# --------------------------------------------------------------------------- #
# Récupération de l'arborescence
# --------------------------------------------------------------------------- #


@dataclass
class Node:
    id: str
    type: str  # page | folder | whiteboard | database | embed
    title: str
    level: int = 0
    path: str = ""
    parent_id: str | None = None
    page: dict | None = None  # détails de la page (corps, auteur, dates…)


@dataclass
class ExportResult:
    root_title: str
    root_type: str
    nodes: list[Node]
    users: dict[str, str]
    space_key: str | None
    notes: list[str] = field(default_factory=list)


_PLURAL = {"page": "pages", "folder": "folders", "whiteboard": "whiteboards", "database": "databases", "embed": "embeds"}


def _sort_children(items):
    return sorted(items, key=lambda c: (c.get("childPosition") is None, c.get("childPosition") or 0))


def walk_tree(client: ConfluenceClient, root: Node, notes: list[str]) -> list[Node]:
    nodes = [root]
    stack = [root]
    while stack:
        parent = stack.pop()
        plural = _PLURAL.get(parent.type)
        if not plural:
            continue
        children = list(
            client.paginate(f"/{plural}/{parent.id}/direct-children", params={"limit": 250}, allow_404=True)
        )
        new = []
        for c in _sort_children(children):
            if c.get("status") not in (None, "current"):
                continue
            node = Node(
                id=str(c.get("id")),
                type=c.get("type") or "page",
                title=c.get("title") or "(sans titre)",
                level=parent.level + 1,
                parent_id=parent.id,
                path=f"{parent.path} / {c.get('title') or '(sans titre)'}",
            )
            new.append(node)
        # Insérer les enfants juste après leur parent pour garder l'ordre de l'arborescence.
        idx = nodes.index(parent)
        nodes[idx + 1 : idx + 1] = new
        stack.extend(reversed(new))
        if len(nodes) > MAX_ITEMS:
            notes.append(f"Export limité aux {MAX_ITEMS} premiers éléments.")
            break
    return nodes


def fetch_pages(client: ConfluenceClient, ids: list[str]) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for i in range(0, len(ids), 100):
        batch = ids[i : i + 100]
        params = [("id", x) for x in batch] + [("body-format", "storage"), ("limit", "250")]
        for p in client.paginate("/pages", params=params):
            out[str(p.get("id"))] = p
    return out


def fetch_users(client: ConfluenceClient, account_ids: set[str]) -> dict[str, str]:
    users: dict[str, str] = {}
    ids = sorted(a for a in account_ids if a)
    for i in range(0, len(ids), 100):
        try:
            r = client.request(
                "POST",
                client.api("/users-bulk"),
                json_body={"accountIds": ids[i : i + 100]},
                allow_404=True,
            )
            if r is None:
                break
            for u in r.json().get("results", []) or []:
                name = u.get("displayName") or u.get("publicName")
                if u.get("accountId") and name:
                    users[u["accountId"]] = name
        except ExportError as e:
            if e.status in (401,):
                raise
            break  # les noms d'auteurs sont un bonus : on continue sans
        except ValueError:
            break
    return users


def resolve_tiny(client: ConfluenceClient, target: Target) -> Target:
    url = f"{target.wiki}/x/{target.id}"
    r = client.request("GET", url, allow_redirects=False)
    hops = 0
    while r is not None and r.status_code in (301, 302, 303, 307, 308) and hops < 5:
        url = urljoin(url, r.headers.get("Location", ""))
        resolved = None
        try:
            resolved = parse_confluence_url(url)
        except ExportError:
            pass
        if resolved and resolved.kind in ("page", "folder", "space"):
            return resolved
        r = client.request("GET", url, allow_redirects=False)
        hops += 1
    raise ExportError("Ce lien court Confluence n'a pas pu être ouvert. Copiez l'URL complète du dossier.")


def collect(client: ConfluenceClient, target: Target) -> ExportResult:
    notes: list[str] = []
    if target.kind == "tiny":
        target = resolve_tiny(client, target)
        client.t = target

    if target.kind == "title":
        space = _get_space(client, target.space_key)
        pages = list(client.paginate("/pages", params={"space-id": space["id"], "title": target.title, "limit": 5}))
        if not pages:
            raise ExportError("Page introuvable dans cet espace : vérifiez l'URL.", 404)
        target = Target(target.origin, target.context, "page", id=str(pages[0]["id"]), space_key=target.space_key)

    if target.kind == "space":
        return collect_space(client, target, notes)

    if target.kind == "folder":
        info = client.get_json(f"/folders/{target.id}")
        root = Node(id=target.id, type="folder", title=info.get("title") or "Dossier", path=info.get("title") or "Dossier")
    else:
        info = client.get_json(f"/pages/{target.id}")
        root = Node(id=target.id, type="page", title=info.get("title") or "Page", path=info.get("title") or "Page")

    nodes = walk_tree(client, root, notes)
    users = attach_pages_and_users(client, nodes)
    return ExportResult(root.title, root.type, nodes, users, target.space_key, notes)


def _get_space(client: ConfluenceClient, key: str) -> dict:
    data = client.get_json("/spaces", params={"keys": key, "limit": 1})
    results = (data or {}).get("results") or []
    if not results:
        raise ExportError("Espace introuvable : vérifiez l'URL ou vos droits d'accès.", 404)
    return results[0]


def collect_space(client: ConfluenceClient, target: Target, notes: list[str]) -> ExportResult:
    space = _get_space(client, target.space_key)
    pages = list(client.paginate(f"/spaces/{space['id']}/pages", params={"limit": 250, "body-format": "storage"}))
    if len(pages) > MAX_ITEMS:
        pages = pages[:MAX_ITEMS]
        notes.append(f"Export limité aux {MAX_ITEMS} premières pages de l'espace.")
    by_id: dict[str, Node] = {}
    for p in pages:
        pid = str(p["id"])
        by_id[pid] = Node(id=pid, type="page", title=p.get("title") or "(sans titre)", parent_id=_s(p.get("parentId")), page=p)

    # Les dossiers ne sont pas listés avec les pages : on les récupère à partir de leurs enfants.
    parent_type = {n.parent_id: (n.page or {}).get("parentType") for n in by_id.values() if n.parent_id}
    pending = [pid for pid, t in parent_type.items() if pid not in by_id and t == "folder"]
    while pending:
        fid = pending.pop()
        if fid in by_id:
            continue
        f = client.get_json(f"/folders/{fid}", allow_404=True)
        if not f:
            continue
        node = Node(id=fid, type="folder", title=f.get("title") or "Dossier", parent_id=_s(f.get("parentId")), page=None)
        node.position = f.get("position")  # type: ignore[attr-defined]
        by_id[fid] = node
        if node.parent_id and node.parent_id not in by_id and f.get("parentType") == "folder":
            pending.append(node.parent_id)

    positions = {str(p["id"]): p.get("position") for p in pages}
    positions.update({n.id: getattr(n, "position", None) for n in by_id.values() if n.type == "folder"})
    children: dict[str | None, list[Node]] = {}
    for n in by_id.values():
        parent = n.parent_id if n.parent_id in by_id else None
        children.setdefault(parent, []).append(n)
    for lst in children.values():
        lst.sort(key=lambda n: (positions.get(n.id) is None, positions.get(n.id) or 0, n.title.lower()))

    ordered: list[Node] = []

    def visit(n: Node, level: int, path: str):
        n.level = level
        n.path = f"{path} / {n.title}" if path else n.title
        ordered.append(n)
        for c in children.get(n.id, []):
            visit(c, level + 1, n.path)

    for r in children.get(None, []):
        visit(r, 0, "")

    name = space.get("name") or target.space_key
    account_ids = set()
    for n in ordered:
        if n.page:
            account_ids |= _page_account_ids(n.page)
    users = fetch_users(client, account_ids)
    return ExportResult(name, "space", ordered, users, target.space_key, notes)


def _s(v):
    return str(v) if v is not None else None


def _page_account_ids(p: dict) -> set[str]:
    ids = {p.get("authorId"), (p.get("version") or {}).get("authorId")}
    body = ((p.get("body") or {}).get("storage") or {}).get("value") or ""
    ids |= set(re.findall(r'ri:account-id="([^"]+)"', body))
    return {i for i in ids if i}


def attach_pages_and_users(client: ConfluenceClient, nodes: list[Node]) -> dict[str, str]:
    page_ids = [n.id for n in nodes if n.type == "page"]
    details = fetch_pages(client, page_ids)
    account_ids: set[str] = set()
    for n in nodes:
        if n.type == "page":
            n.page = details.get(n.id)
            if n.page:
                account_ids |= _page_account_ids(n.page)
    return fetch_users(client, account_ids)


# --------------------------------------------------------------------------- #
# Lecture du contenu des pages (format "storage" de Confluence)
# --------------------------------------------------------------------------- #

_BLOCK_TAGS = {
    "p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "pre", "blockquote",
    "ac:task", "ul", "ol", "table", "hr",
}


def _prepare(soup: BeautifulSoup, users: dict[str, str], mention_prefix: str = "@"):
    """Remplace les éléments propres à Confluence par du texte lisible."""
    # Paramètres de macros : invisibles à l'écran, sauf pour le badge de statut.
    for macro in soup.find_all("ac:structured-macro"):
        name = (macro.get("ac:name") or "").lower()
        if name == "status":
            title = macro.find("ac:parameter", attrs={"ac:name": "title"})
            macro.replace_with(NavigableString(title.get_text(strip=True) if title else ""))
        elif name in ("toc", "children", "pagetree", "recently-updated", "contentbylabel", "attachments", "jira"):
            macro.decompose()
    for p in soup.find_all(["ac:parameter", "ac:placeholder"]):
        p.decompose()
    for user in soup.find_all("ri:user"):
        name = users.get(user.get("ri:account-id") or "", "")
        user.replace_with(NavigableString(f"{mention_prefix}{name}" if name else f"{mention_prefix}utilisateur"))
    for t in soup.find_all("time"):
        t.replace_with(NavigableString(_fr_date(t.get("datetime") or "") or t.get_text()))
    for e in soup.find_all("ac:emoticon"):
        e.replace_with(NavigableString(e.get("ac:emoji-fallback") or ""))
    for link in soup.find_all("ac:link"):
        body = link.find(["ac:link-body", "ac:plain-text-link-body"])
        ref = link.find(["ri:page", "ri:attachment", "ri:space"])
        text = body.get_text() if body else ""
        if not text and ref is not None:
            text = ref.get("ri:content-title") or ref.get("ri:filename") or ref.get("ri:space-key") or ""
        if not text:
            text = link.get_text()  # ex. mention @utilisateur déjà convertie
        link.replace_with(NavigableString(text))
    for img in soup.find_all("ac:image"):
        img.replace_with(NavigableString("[image]"))
    for task in soup.find_all("ac:task"):
        status = task.find("ac:task-status")
        body = task.find("ac:task-body")
        done = status is not None and status.get_text(strip=True) == "complete"
        task.replace_with(NavigableString(("☑ " if done else "☐ ") + (body.get_text(" ", strip=True) if body else "")))
    for br in soup.find_all("br"):
        br.replace_with(NavigableString("\n"))


def _fr_date(iso: str) -> str:
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", iso or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else ""


def _normalize(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _render(node, chunks: list[str]):
    if isinstance(node, NavigableString):
        parent = node.parent.name if node.parent is not None else ""
        if parent in ("tr", "tbody", "thead", "tfoot", "table", "colgroup") and not str(node).strip():
            return  # retours à la ligne du code source entre les cellules
        chunks.append(str(node))
        return
    if not isinstance(node, Tag):
        return
    if node.name in ("td", "th"):
        # Dans le texte de la page, une cellule tient sur une ligne : « a | b | c »
        chunks.append(" " + _cell_text(node).replace("\n", " ") + " |")
        return
    block = node.name in _BLOCK_TAGS
    if block:
        chunks.append("\n")
    for c in node.children:
        _render(c, chunks)
    if block:
        chunks.append("\n")


def _text_of(el) -> str:
    chunks: list[str] = []
    _render(el, chunks)
    return _normalize("".join(chunks))


def _cell_text(cell) -> str:
    chunks: list[str] = []
    for c in cell.children:
        _render(c, chunks)
    return _normalize("".join(chunks))


@dataclass
class TableData:
    page_title: str
    page_url: str
    rows: list[list[str]]
    header: bool


def parse_body(storage: str, users: dict[str, str], page_title: str, page_url: str):
    if not storage:
        return "", []
    soup = BeautifulSoup(storage, "html.parser")
    _prepare(soup, users)
    tables: list[TableData] = []
    for table in soup.find_all("table"):
        if table.find_parent("table") is not None:
            continue  # tableau imbriqué : son texte reste dans la cellule parente
        rows, header = _table_rows(table)
        if rows and any(any(c for c in r) for r in rows):
            tables.append(TableData(page_title, page_url, rows, header))
    return _text_of(soup), tables


def _table_rows(table: Tag):
    grid: list[list[str]] = []
    spans: dict[int, tuple[int, str]] = {}  # colonne -> (lignes restantes, valeur)
    first_row_is_header = False
    trs = [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]
    for r_index, tr in enumerate(trs):
        cells = [c for c in tr.find_all(["td", "th"], recursive=False)]
        row: list[str] = []
        col = 0

        def fill_spans():
            nonlocal col
            while col in spans:
                left, val = spans[col]
                row.append("")
                if left <= 1:
                    del spans[col]
                else:
                    spans[col] = (left - 1, val)
                col += 1

        for cell in cells:
            fill_spans()
            text = _cell_text(cell)
            colspan = _int(cell.get("colspan"), 1)
            rowspan = _int(cell.get("rowspan"), 1)
            for k in range(colspan):
                row.append(text if k == 0 else "")
                if rowspan > 1:
                    spans[col] = (rowspan - 1, text)
                col += 1
        fill_spans()
        if r_index == 0 and cells and (all(c.name == "th" for c in cells) or all(_is_bold(c) for c in cells)):
            first_row_is_header = True
        grid.append(row)
    width = max((len(r) for r in grid), default=0)
    grid = [r + [""] * (width - len(r)) for r in grid]
    return grid, first_row_is_header


def _is_bold(cell) -> bool:
    """Cellule dont tout le texte est en gras : utilisée comme en-tête."""
    text = cell.get_text(strip=True)
    if not text:
        return False
    bold = "".join(b.get_text(strip=True) for b in cell.find_all(["strong", "b"]) if not b.find_parent(["strong", "b"]))
    return bold == text


def _int(v, default):
    try:
        return max(1, min(int(v), 50))
    except (TypeError, ValueError):
        return default


# --------------------------------------------------------------------------- #
# Construction du fichier Excel
# --------------------------------------------------------------------------- #

_HEAD_FILL = PatternFill("solid", fgColor=GREEN)
_HEAD_FONT = Font(bold=True, color="FFFFFF")
_THIN = Side(style="thin", color="D6E3C0")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TYPE_FR = {"page": "Page", "folder": "Dossier", "whiteboard": "Tableau blanc", "database": "Base de données", "embed": "Lien intégré", "space": "Espace"}
_NUMBER_RE = re.compile(r"^-?(0|[1-9]\d{0,2}(?:[ \u202f\xa0]\d{3})*|[1-9]\d*)(?:[.,]\d+)?$")


def _clean(value):
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
        if len(value) > EXCEL_CELL_LIMIT:
            value = value[: EXCEL_CELL_LIMIT - 1] + "…"
    return value


def _as_text(cell):
    """Un texte commençant par « = » reste du texte : jamais exécuté comme formule."""
    if isinstance(cell.value, str) and cell.value.startswith("="):
        cell.data_type = "s"
    return cell


def _maybe_number(text: str):
    t = text.strip()
    if t and _NUMBER_RE.match(t):
        try:
            return float(re.sub(r"[ \u202f\xa0]", "", t).replace(",", ".")) if re.search(r"[.,]", t) else int(re.sub(r"[ \u202f\xa0]", "", t))
        except ValueError:
            return text
    m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", t)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return text
    return text


def _to_local(iso: str | None):
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(PARIS).replace(tzinfo=None)
    except ValueError:
        return None


def _sheet_name(base: str, used: set[str]) -> str:
    name = re.sub(r"[\[\]\:\*\?\/\\]", " ", base).strip().strip("'") or "Feuille"
    name = name[:31]
    candidate, n = name, 2
    while candidate.lower() in used:
        suffix = f" ({n})"
        candidate = name[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def _style_header(ws, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = _HEAD_FILL
        cell.font = _HEAD_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = _BORDER
    ws.row_dimensions[row].height = 22


def _autosize(ws, min_w=8, max_w=60, start_row=1):
    widths: dict[int, int] = {}
    for row in ws.iter_rows(min_row=start_row):
        for cell in row:
            if cell.value is None:
                continue
            v = cell.value
            text = v.strftime("%d/%m/%Y %H:%M") if isinstance(v, datetime) else str(v)
            longest = max((len(line) for line in text.split("\n")), default=0)
            widths[cell.column] = max(widths.get(cell.column, 0), longest)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = max(min_w, min(max_w, w + 2))


def build_workbook(res: ExportResult, target: Target, include_content: bool = True):
    wb = Workbook()
    used: set[str] = set()
    summary = wb.active
    summary.title = _sheet_name("Résumé", used)
    pages_ws = wb.create_sheet(_sheet_name("Pages", used))

    # ---- Feuille "Pages" ----------------------------------------------------
    headers = ["Niveau", "Titre", "Type", "Chemin", "Créé par", "Créé le", "Modifié par", "Modifié le", "Version", "Lien"]
    if include_content:
        headers.append("Contenu")
    pages_ws.append(headers)
    _style_header(pages_ws, 1, len(headers))

    all_tables: list[TableData] = []
    n_pages = n_folders = 0
    for node in res.nodes:
        p = node.page or {}
        links = p.get("_links") or {}
        if links.get("webui"):
            url = target.wiki + links["webui"]
        elif node.type == "folder" and res.space_key:
            url = f"{target.wiki}/spaces/{res.space_key}/folder/{node.id}"
        elif node.type == "page":
            url = f"{target.wiki}/pages/viewpage.action?pageId={node.id}"
        else:
            url = ""
        storage = ((p.get("body") or {}).get("storage") or {}).get("value") or ""
        text, tables = parse_body(storage, res.users, node.title, url) if node.type == "page" else ("", [])
        all_tables.extend(tables)
        if node.type == "page":
            n_pages += 1
        elif node.type == "folder":
            n_folders += 1
        version = p.get("version") or {}
        row = [
            node.level,
            node.title,
            _TYPE_FR.get(node.type, node.type),
            node.path,
            res.users.get(p.get("authorId") or "", ""),
            _to_local(p.get("createdAt")),
            res.users.get(version.get("authorId") or "", ""),
            _to_local(version.get("createdAt")),
            version.get("number"),
            "Ouvrir" if url else "",
        ]
        if include_content:
            row.append(text)
        pages_ws.append([_clean(v) for v in row])
        r = pages_ws.max_row
        for c in pages_ws[r]:
            _as_text(c)
        title_cell = pages_ws.cell(row=r, column=2)
        title_cell.alignment = Alignment(indent=min(node.level * 2, 15))
        if node.type == "folder":
            title_cell.font = Font(bold=True, color=GREEN_DARK)
        if url:
            link = pages_ws.cell(row=r, column=10)
            link.hyperlink = url
            link.font = Font(color=GREEN_DARK, underline="single")
        for col in (6, 8):
            pages_ws.cell(row=r, column=col).number_format = "DD/MM/YYYY HH:MM"
        if include_content:
            pages_ws.cell(row=r, column=11).alignment = Alignment(vertical="top", wrap_text=False)
        pages_ws.row_dimensions[r].height = 18  # une ligne par page, même si le texte est long

    _autosize(pages_ws)
    pages_ws.column_dimensions["A"].width = 8
    pages_ws.column_dimensions["B"].width = 45
    pages_ws.column_dimensions["D"].width = 50
    if include_content:
        pages_ws.column_dimensions["K"].width = 90
    pages_ws.freeze_panes = "C2"
    pages_ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(pages_ws.max_row, 1)}"

    # ---- Une feuille par tableau trouvé ------------------------------------
    table_sheets: list[tuple[str, TableData]] = []
    for i, tbl in enumerate(all_tables[:MAX_TABLE_SHEETS], start=1):
        name = _sheet_name(f"T{i} - {tbl.page_title}", used)
        ws = wb.create_sheet(name)
        ws["A1"] = _clean(f"Tableau {i} - page « {tbl.page_title} »")
        ws["A1"].font = Font(bold=True, size=12, color=GREEN_DARK)
        if tbl.page_url:
            ws["A1"].hyperlink = tbl.page_url
        start = 3
        for r_off, row in enumerate(tbl.rows):
            for c_off, val in enumerate(row):
                v = val if (tbl.header and r_off == 0) else _maybe_number(val)
                cell = _as_text(ws.cell(row=start + r_off, column=1 + c_off, value=_clean(v)))
                cell.border = _BORDER
                cell.alignment = Alignment(vertical="top", wrap_text="\n" in str(val))
                if isinstance(v, datetime):
                    cell.number_format = "DD/MM/YYYY"
        ncols = max((len(r) for r in tbl.rows), default=1)
        last_row = start + len(tbl.rows) - 1
        if tbl.header:
            _style_header(ws, start, ncols)
            ws.freeze_panes = ws.cell(row=start + 1, column=1)
            if len(tbl.rows) > 1:
                ws.auto_filter.ref = f"A{start}:{get_column_letter(ncols)}{last_row}"
        _autosize(ws, start_row=start)
        table_sheets.append((name, tbl))

    # ---- Feuille "Résumé" ---------------------------------------------------
    s = summary
    s["A1"] = "Export Confluence"
    s["A1"].font = Font(bold=True, size=18, color=GREEN_DARK)
    now = datetime.now(PARIS).replace(tzinfo=None) if PARIS is not timezone.utc else datetime.utcnow()
    info = [
        ("Source", _TYPE_FR.get(res.root_type, res.root_type) + " « " + res.root_title + " »"),
        ("Lien", target.original_url),
        ("Exporté le", now),
        ("Nombre de pages", n_pages),
        ("Nombre de dossiers", n_folders),
        ("Nombre de tableaux", len(all_tables)),
    ]
    for i, (k, v) in enumerate(info, start=3):
        s.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = _as_text(s.cell(row=i, column=2, value=_clean(v)))
        c.alignment = Alignment(horizontal="left")
        if k == "Lien":
            c.hyperlink = target.original_url
            c.font = Font(color=GREEN_DARK, underline="single")
        if k == "Exporté le":
            c.number_format = "DD/MM/YYYY HH:MM"
    row = 3 + len(info) + 1
    notes = list(res.notes)
    if len(all_tables) > MAX_TABLE_SHEETS:
        notes.append(f"{len(all_tables)} tableaux trouvés : seuls les {MAX_TABLE_SHEETS} premiers ont une feuille.")
    if not include_content:
        notes.append("Le texte des pages a été retiré pour que le fichier reste téléchargeable.")
    for n in notes:
        s.cell(row=row, column=1, value="Note").font = Font(bold=True, color="B45309")
        s.cell(row=row, column=2, value=_clean(n))
        row += 1
    row += 1
    s.cell(row=row, column=1, value="Feuilles").font = Font(bold=True)
    row += 1
    for name, desc in [(pages_ws.title, "Arborescence complète avec auteurs, dates et liens")] + [
        (n, f"Tableau de la page « {t.page_title} »") for n, t in table_sheets
    ]:
        c = s.cell(row=row, column=1, value=name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = Font(color=GREEN_DARK, underline="single")
        s.cell(row=row, column=2, value=_clean(desc))
        row += 1
    s.column_dimensions["A"].width = 28
    s.column_dimensions["B"].width = 80
    s.sheet_view.showGridLines = False

    buf = io.BytesIO()
    wb.save(buf)
    stats = {"pages": n_pages, "folders": n_folders, "tables": len(all_tables)}
    return buf.getvalue(), stats


# --------------------------------------------------------------------------- #
# Mode « fiches applicatives » : reproduit Liste_des_Applications_CNSA.xlsx
#   - une feuille par sous-dossier du dossier source (les pôles)
#   - une ligne par fiche applicative, une colonne par champ de la fiche
# --------------------------------------------------------------------------- #

FICHE_COLUMNS = [
    "Nom", "Code", "Version", "Typologie", "Description de l'application", "Description fonctionnelle",
    "Fonction principale", "Enjeux et contexte", "Année de mise en service", "Propriétaire SI", "Propriétaire",
    "Responsable DSI", "Utilisateurs internes", "Utilisateurs externes", "Thématique SDSI",
    "Lien avec Engagement COG", "Impact sur le SI", "Interfaces internes", "Interfaces externes",
    "Zone fonctionnelle", "Quartier fonctionnel", "Îlot fonctionnel", "Bloc fonctionnel", "Périmètre fonctionnel",
    "Champs de données", "État actuel", "Technologies utilisées", "Consommation du service", "Hébergement",
    "Flux de données", "Performance", "Niveau de disponibilité", "Volumétrie de l'application",
    "Plan de sauvegarde", "Fréquence de mise à jour", "Cadres applicables", "Chiffrement", "Stockage", "Coûts",
    "Plan de remplacement", "Date de décommissionnement",
]
FICHE_WIDTHS = [32, 20, 12, 14, 46, 46, 40, 46, 14] + [24] * 16 + [16] + [24] * 15

# Autres intitulés possibles dans les fiches -> colonne de l'Excel
FICHE_ALIASES = {
    "nom de l application": "Nom", "nom application": "Nom", "application": "Nom", "nom de l outil": "Nom",
    "code application": "Code", "code de l application": "Code", "identifiant": "Code",
    "version actuelle": "Version", "type": "Typologie", "type d application": "Typologie",
    "description": "Description de l'application", "description application": "Description de l'application",
    "fonctionnalites": "Description fonctionnelle", "fonction": "Fonction principale",
    "enjeux": "Enjeux et contexte", "contexte": "Enjeux et contexte",
    "annee de mise en production": "Année de mise en service", "date de mise en service": "Année de mise en service",
    "mise en service": "Année de mise en service",
    "proprietaire metier": "Propriétaire", "responsable": "Responsable DSI", "referent dsi": "Responsable DSI",
    "thematique": "Thématique SDSI", "engagement cog": "Lien avec Engagement COG",
    "etat": "État actuel", "statut": "État actuel", "etat de l application": "État actuel",
    "technologies": "Technologies utilisées", "technologie": "Technologies utilisées",
    "technologie utilisee": "Technologies utilisées", "hebergeur": "Hébergement",
    "disponibilite": "Niveau de disponibilité", "volumetrie": "Volumétrie de l'application",
    "sauvegarde": "Plan de sauvegarde", "frequence de mise a jour": "Fréquence de mise à jour",
    "cadre applicable": "Cadres applicables", "cout": "Coûts", "couts": "Coûts",
    "date de decommissionnement prevue": "Date de décommissionnement", "decommissionnement": "Date de décommissionnement",
}

_FICHE_HEAD_FILL = PatternFill("solid", fgColor="5E9412")  # vert CNSA foncé (texte blanc lisible)
_FICHE_ZEBRA = PatternFill("solid", fgColor="EEF6E0")  # vert très clair une ligne sur deux
_FICHE_SIDE = Side(style="thin", color="C5DFA0")
_FICHE_BORDER = Border(left=_FICHE_SIDE, right=_FICHE_SIDE, top=_FICHE_SIDE, bottom=_FICHE_SIDE)


def _norm_label(text: str) -> str:
    t = unicodedata.normalize("NFKD", text or "")
    t = "".join(c for c in t if not unicodedata.combining(c)).lower()
    t = t.replace("œ", "oe").replace("’", " ").replace("'", " ")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


_COL_BY_NORM = {_norm_label(c): c for c in FICHE_COLUMNS}
_COL_BY_NORM.update({_norm_label(k): v for k, v in FICHE_ALIASES.items()})


def match_label(text: str) -> str | None:
    """Associe un intitulé de la fiche à une colonne de l'Excel (tolère accents, casse, « : », fautes légères)."""
    n = _norm_label(text)
    if not n or len(n) > 60:
        return None
    if n in _COL_BY_NORM:
        return _COL_BY_NORM[n]
    import difflib

    close = difflib.get_close_matches(n, list(_COL_BY_NORM), n=1, cutoff=0.9)
    return _COL_BY_NORM[close[0]] if close else None


def _clean_label(text: str) -> str:
    """Intitulé tel que le script d'origine le lit : sans « * » ni « : », espaces simplifiés."""
    return " ".join((text or "").replace("*", "").replace(":", "").split())


def _value_text(el) -> str:
    """Texte d'une valeur, sur une seule ligne (les listes sont séparées par des virgules)."""
    if isinstance(el, NavigableString):
        return re.sub(r"\s+", " ", str(el)).strip()
    text = _cell_text(el)
    lines = [l.strip(" \t•·-–") if el.find("li") else l.strip() for l in text.split("\n")]
    lines = [l for l in lines if l]
    return (", " if el.find("li") else " ").join(lines).strip()


@dataclass
class FicheParse:
    fields: dict
    unknown: list  # intitulés non reconnus (pour la feuille Contrôle)


def parse_fiche(storage: str, users: dict[str, str]) -> FicheParse:
    fields: dict[str, str] = {}
    unknown: list[str] = []
    if not storage:
        return FicheParse(fields, unknown)
    soup = BeautifulSoup(storage, "html.parser")
    _prepare(soup, users, mention_prefix="")

    def put(col, value):
        value = (value or "").strip()
        if value.strip(" |").lower().startswith("non renseign"):
            value = ""  # valeur par défaut du modèle de fiche
        if value and _norm_label(value) == _norm_label(col):
            return  # la cellule répète l'intitulé
        if value and not fields.get(col):
            fields[col] = value
        elif col not in fields:
            fields[col] = value

    # 1) Tableaux « Intitulé | Valeur » : c'est la forme des fiches CNSA (règle du script d'origine :
    #    une ligne de 2 cellules = un champ ; un intitulé inconnu devient une nouvelle colonne).
    #    Sont aussi acceptés : 4 colonnes « Intitulé | Valeur | Intitulé | Valeur » et tableau horizontal.
    for table in soup.find_all("table"):
        if table.find_parent("table") is not None:
            continue
        trs = [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]
        grid = [[c for c in tr.find_all(["td", "th"], recursive=False)] for tr in trs]
        if not grid:
            continue
        head = grid[0]
        head_cols = [match_label(_value_text(c)) for c in head]
        vertical_hits = sum(1 for row in grid if row and match_label(_value_text(row[0])))
        if len(grid) >= 2 and sum(1 for c in head_cols if c) >= max(2, len(head) // 2) and sum(
            1 for c in head_cols if c
        ) > vertical_hits:
            for i, col in enumerate(head_cols):
                if col and i < len(grid[1]):
                    put(col, _value_text(grid[1][i]))
            continue
        for row in grid:
            if len(row) == 2:
                label = _clean_label(_value_text(row[0]))
                if label and len(label) <= 80:
                    put(match_label(label) or label, _value_text(row[1]))
                continue
            i = 0
            while i < len(row):
                col = match_label(_value_text(row[i]))
                if col and i + 1 < len(row):
                    put(col, _value_text(row[i + 1]))
                    i += 2
                    continue
                i += 1

    # 2) Titres (« Description fonctionnelle » en titre, puis le texte en dessous)
    def starts_with_label(el) -> bool:
        txt = _value_text(el)
        m = re.match(r"^\s*([^:|]{2,60}?)\s*:", txt)
        if m and match_label(m.group(1)):
            return True
        first = el.find(["strong", "b"]) if isinstance(el, Tag) else None
        return bool(first and match_label(first.get_text(" ", strip=True)) and txt.startswith(first.get_text(strip=True)[:3]))

    for h in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6"]):
        col = match_label(h.get_text(" ", strip=True))
        if not col or col in fields:
            continue
        parts = []
        for sib in h.find_next_siblings():
            if sib.name in ("h1", "h2", "h3", "h4", "h5", "h6", "table", "ac:structured-macro") or starts_with_label(sib):
                break
            parts.append(_value_text(sib))
        put(col, " ".join(p for p in parts if p))

    # 3) Lignes « Intitulé : valeur »
    for line in _text_of(soup).split("\n"):
        m = re.match(r"^\s*([^:|]{2,60}?)\s*:\s*(.+?)\s*$", line)
        if m:
            col = match_label(m.group(1))
            if col and not fields.get(col) and "|" not in m.group(2):  # « | » : ligne de tableau déjà lue
                put(col, m.group(2).lstrip(": ").strip())
    return FicheParse(fields, unknown)


@dataclass
class FicheSheet:
    title: str
    rows: list  # liste de dict colonne -> valeur


@dataclass
class FicheResult:
    root_title: str
    sheets: list
    ignored: list  # (pôle, titre de page, lien, nb champs reconnus)
    unknown: dict  # intitulé -> (nombre, exemple de page)
    notes: list


def fetch_tree(client: ConfluenceClient, root_id: str, root_type: str, notes: list[str]) -> list[Node]:
    """Toute l'arborescence sous la racine, en une seule requête quand c'est possible."""
    root = Node(id=root_id, type=root_type, title="", level=0)
    plural = _PLURAL.get(root_type, "pages")
    try:
        items = list(client.paginate(f"/{plural}/{root_id}/descendants", params={"limit": 250, "depth": 5}))
    except ExportError as e:
        if e.status in (401, 403, 413):
            raise
        items = None
    if items is None:
        return walk_tree(client, root, notes)[1:]
    by_parent: dict[str, list[dict]] = {}
    for it in items:
        if it.get("status") not in (None, "current"):
            continue
        by_parent.setdefault(str(it.get("parentId")), []).append(it)
    out: list[Node] = []

    def visit(pid: str, level: int):
        for c in _sort_children(by_parent.get(pid, [])):
            n = Node(id=str(c["id"]), type=c.get("type") or "page", title=c.get("title") or "(sans titre)", level=level, parent_id=pid)
            out.append(n)
            if level >= 5 and n.type in _PLURAL:
                deeper = walk_tree(client, Node(id=n.id, type=n.type, title=n.title, level=level), notes)[1:]
                out.extend(deeper)
            else:
                visit(n.id, level + 1)

    visit(root_id, 1)
    return out


def collect_fiches(client: ConfluenceClient, target: Target) -> FicheResult:
    notes: list[str] = []
    if target.kind == "tiny":
        target = resolve_tiny(client, target)
        client.t = target
    if target.kind not in ("page", "folder"):
        raise ExportError("Pour la liste des applications, indiquez le lien du dossier « Fiche applicative ».")
    info = client.get_json(f"/{'folders' if target.kind == 'folder' else 'pages'}/{target.id}")
    root_title = info.get("title") or "Applications"
    nodes = fetch_tree(client, target.id, target.kind, notes)
    if len(nodes) > MAX_ITEMS:
        nodes = nodes[:MAX_ITEMS]
        notes.append(f"Export limité aux {MAX_ITEMS} premières pages.")
    users = attach_pages_and_users(client, nodes)

    children: dict[str, list[Node]] = {}
    for n in nodes:
        children.setdefault(n.parent_id, []).append(n)

    def descendants(nid):
        for c in children.get(nid, []):
            yield c
            yield from descendants(c.id)

    groups: list[tuple[str, list[Node]]] = []
    loose: list[Node] = []
    for top in children.get(target.id, []):
        if children.get(top.id):
            groups.append((top.title, list(descendants(top.id))))
        else:
            loose.append(top)
    if loose:
        groups.append(("Racine", loose))

    sheets, ignored, unknown, extra_cols = [], [], {}, []
    for title, pages in sorted(groups, key=lambda g: g[0]):
        rows = []
        for n in pages:
            if n.type != "page" or not n.page:
                continue
            storage = ((n.page.get("body") or {}).get("storage") or {}).get("value") or ""
            parsed = parse_fiche(storage, users)
            if len(parsed.fields) < 2:  # page sans tableau de fiche : dossier, sommaire…
                continue
            for k in parsed.fields:
                if k not in FICHE_COLUMNS and k not in extra_cols:
                    extra_cols.append(k)
            row = dict(parsed.fields)
            if not row.get("Nom"):
                row["Nom"] = n.title
            rows.append(row)
        if rows:
            rows.sort(key=lambda r: str(r.get("Nom", "")).lower())  # même ordre que l'Excel d'origine
            sheets.append(FicheSheet(title, rows))
    for k in extra_cols:
        unknown[k] = (sum(1 for sh in sheets for r in sh.rows if k in r), "")
    return FicheResult(root_title, sheets, ignored, unknown, notes)


def build_fiches_workbook(res: FicheResult):
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    head_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font = Font(name="Calibri", size=10)
    bold = Font(name="Calibri", size=10, bold=True)
    empty_font = Font(name="Calibri", size=10, color="9BA3AB", italic=True)  # cases vides, comme l'original
    head_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_align = Alignment(vertical="center", indent=1)
    columns = FICHE_COLUMNS + [k for k in res.unknown if k not in FICHE_COLUMNS]  # nouveaux champs à la fin
    last_col = get_column_letter(len(columns))
    total = 0
    for sh in res.sheets or [FicheSheet(res.root_title, [])]:
        ws = wb.create_sheet(_sheet_name(sh.title, used))
        ws.append(columns)
        for c in ws[1]:
            c.font, c.fill, c.alignment, c.border = head_font, _FICHE_HEAD_FILL, head_align, _FICHE_BORDER
        ws.row_dimensions[1].height = 32
        for i, row in enumerate(sh.rows, start=2):
            for j, col in enumerate(columns, start=1):
                v = row.get(col)
                cell = _as_text(ws.cell(row=i, column=j, value=_clean(v) if v else None))
                cell.font = bold if j == 1 else (font if v else empty_font)
                cell.alignment = cell_align
                cell.border = _FICHE_BORDER
                if i % 2 == 1:
                    cell.fill = _FICHE_ZEBRA
            ws.row_dimensions[i].height = 17
        for j in range(1, len(columns) + 1):
            ws.column_dimensions[get_column_letter(j)].width = FICHE_WIDTHS[j - 1] if j <= len(FICHE_WIDTHS) else 24
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{last_col}{max(ws.max_row, 1)}"
        total += len(sh.rows)

    # Feuille de contrôle, uniquement s'il y a quelque chose à vérifier
    if res.notes or total == 0:
        ws = wb.create_sheet(_sheet_name("Contrôle", used))
        ws.sheet_properties.tabColor = "C00000"
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 50
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 60
        r = 1
        ws.cell(r, 1, "Points à vérifier").font = Font(bold=True, size=13, color="C00000")
        r += 2
        msgs = list(res.notes)
        if total == 0:
            msgs.append("Aucune fiche reconnue : envoyez ce fichier à la personne qui maintient l'outil.")
        for m in msgs:
            ws.cell(r, 1, _clean(m)); r += 1
    buf = io.BytesIO()
    wb.save(buf)
    stats = {"pages": total, "folders": len(res.sheets), "tables": 0, "applications": total, "poles": len(res.sheets)}
    if total == 0:
        stats["note"] = "Aucune fiche reconnue : voir la feuille Contrôle"
    return buf.getvalue(), stats


# --------------------------------------------------------------------------- #
# Point d'entrée principal
# --------------------------------------------------------------------------- #


def safe_filename(name: str) -> str:
    name = unicodedata.normalize("NFC", name or "").strip()
    name = re.sub(r"\.xlsx$", "", name, flags=re.I)
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', " ", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return (name[:120] or "export_confluence") + ".xlsx"


def run_export(url: str, email: str, token: str, filename: str | None = None, fmt: str | None = None):
    """Retourne (contenu_xlsx, nom_du_fichier, statistiques)."""
    email = (email or "").strip()
    token = (token or "").strip()
    if not email or "@" not in email:
        raise ExportError("Renseignez l'email de votre compte Atlassian.")
    if not token:
        raise ExportError("Renseignez votre token API Atlassian.")
    target = parse_confluence_url(url)
    deadline = time.monotonic() + MAX_SECONDS
    client = ConfluenceClient(target, email, token, deadline)
    fmt = (fmt or EXPORT_FORMAT or "fiches").lower()

    if fmt == "fiches":
        res = collect_fiches(client, target)
        content, stats = build_fiches_workbook(res)
        stats["title"] = res.root_title
        base = re.sub(r"\.xlsx$", "", (filename or DEFAULT_FILENAME or res.root_title).strip(), flags=re.I)
        return content, safe_filename(f"{base} - {horodatage()}"), stats

    res = collect(client, target)
    target = client.t  # le lien court a pu être remplacé par l'URL réelle
    target.original_url = parse_confluence_url(url).original_url
    content, stats = build_workbook(res, target, include_content=True)
    if len(content) > MAX_RESPONSE_BYTES:
        content, stats = build_workbook(res, target, include_content=False)
        stats["note"] = "Le texte des pages a été retiré pour alléger le fichier."
    if len(content) > MAX_RESPONSE_BYTES:
        raise ExportError(
            "Le fichier Excel dépasse la taille maximale autorisée en ligne. "
            "Exportez un sous-dossier plus petit.",
            413,
        )
    stats["title"] = res.root_title
    return content, safe_filename(filename or res.root_title), stats


def horodatage() -> str:
    """Date et heure de Paris pour le nom du fichier, ex. « 2026-10-20 14h05 » : chaque export a son propre nom."""
    now = datetime.now(PARIS) if PARIS is not timezone.utc else datetime.now(timezone.utc)
    return now.strftime("%Y-%m-%d %Hh%M")


def server_configured() -> bool:
    return bool(CONF_EMAIL and CONF_TOKEN)


def check_access_code(code: str):
    import hmac

    if ACCESS_CODE and not hmac.compare_digest((code or "").strip().encode(), ACCESS_CODE.encode()):
        time.sleep(1.0)  # ralentit les essais au hasard
        raise ExportError("Code d'accès incorrect.", 403)


# --------------------------------------------------------------------------- #
# Handler HTTP (format attendu par Vercel)
# --------------------------------------------------------------------------- #


class handler(BaseHTTPRequestHandler):
    server_version = "ExportConfluence/1.0"

    def log_message(self, format, *args):  # noqa: A002 - pas de logs (identifiants)
        return

    def _json(self, status: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path.split("?")[0].rstrip("/").endswith("/api/export"):
            return self._json(
                200,
                {
                    "ok": True,
                    "configured": server_configured(),
                    "needsCode": bool(ACCESS_CODE),
                    "defaultUrl": SOURCE_URL,
                    "defaultFilename": DEFAULT_FILENAME,
                },
            )
        return self._json(404, {"error": "Page introuvable."})

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length <= 0 or length > 20_000:
                raise ExportError("Requête invalide.")
            try:
                data = json.loads(self.rfile.read(length).decode("utf-8"))
            except (ValueError, UnicodeDecodeError):
                raise ExportError("Requête invalide.")
            if not isinstance(data, dict):
                raise ExportError("Requête invalide.")
            check_access_code(str(data.get("code") or ""))
            if server_configured():
                # Dossier, compte et token fixés côté serveur : rien ne vient du visiteur.
                content, fname, stats = run_export(SOURCE_URL, CONF_EMAIL, CONF_TOKEN, DEFAULT_FILENAME)
            else:
                content, fname, stats = run_export(
                    str(data.get("url") or ""),
                    str(data.get("email") or ""),
                    str(data.get("token") or ""),
                    str(data.get("filename") or "") or None,
                )
        except ExportError as e:
            return self._json(e.status, {"error": e.message})
        except Exception:  # erreur imprévue : message générique, sans détails sensibles
            return self._json(500, {"error": "Une erreur inattendue est survenue pendant l'export. Réessayez."})

        ascii_name = unicodedata.normalize("NFKD", fname).encode("ascii", "ignore").decode() or "export.xlsx"
        ascii_name = ascii_name.replace('"', "")
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Length", str(len(content)))
        self.send_header(
            "Content-Disposition", f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(fname)}"
        )
        self.send_header("X-Export-Filename", quote(fname))
        self.send_header("X-Export-Pages", str(stats["pages"]))
        self.send_header("X-Export-Folders", str(stats["folders"]))
        self.send_header("X-Export-Tables", str(stats["tables"]))
        if "applications" in stats:
            self.send_header("X-Export-Applications", str(stats["applications"]))
            self.send_header("X-Export-Poles", str(stats["poles"]))
        if stats.get("note"):
            self.send_header("X-Export-Note", quote(stats["note"]))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(content)
