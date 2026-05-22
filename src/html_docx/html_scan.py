from __future__ import annotations

from dataclasses import dataclass, field
from html.parser import HTMLParser


@dataclass
class HTMLNodeSnapshot:
    id: str
    tag: str
    attrs: dict[str, str]
    text_parts: list[str] = field(default_factory=list)
    chunks: list["HTMLRunChunk"] = field(default_factory=list)
    has_segments: bool = False

    @property
    def kind(self) -> str | None:
        return self.attrs.get("data-hdocx-type")

    @property
    def lock(self) -> str | None:
        return self.attrs.get("data-hdocx-lock")

    @property
    def text(self) -> str:
        return "".join(self.text_parts)

    def append_run_text(self, data: str, segment_attrs: dict[str, str] | None) -> None:
        self.text_parts.append(data)
        attrs = segment_attrs or {}
        if segment_attrs is not None:
            self.has_segments = True
        if self.chunks and self.chunks[-1].attrs == attrs:
            self.chunks[-1].text_parts.append(data)
        else:
            self.chunks.append(HTMLRunChunk(dict(attrs), [data]))


@dataclass
class HTMLRunChunk:
    attrs: dict[str, str]
    text_parts: list[str]

    @property
    def text(self) -> str:
        return "".join(self.text_parts)


class HDocxHTMLScanner(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.ids: list[str] = []
        self.nodes: dict[str, HTMLNodeSnapshot] = {}
        self.stack: list[tuple[str, str | None, dict[str, str] | None]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = {name: value or "" for name, value in attrs}
        node_id = attr_map.get("data-hdocx-id")
        if node_id:
            self.ids.append(node_id)
            self.nodes[node_id] = HTMLNodeSnapshot(node_id, tag, attr_map)
        segment_attrs = attr_map if attr_map.get("data-hdocx-type") == "run-segment" else None
        self.stack.append((tag, node_id, segment_attrs))

    def handle_endtag(self, tag: str) -> None:
        for index in range(len(self.stack) - 1, -1, -1):
            stacked_tag, _, _ = self.stack[index]
            if stacked_tag == tag:
                del self.stack[index:]
                return

    def handle_data(self, data: str) -> None:
        active_segment_attrs: dict[str, str] | None = None
        for _, _, segment_attrs in reversed(self.stack):
            if segment_attrs is not None:
                active_segment_attrs = segment_attrs
                break
        for _, node_id, _ in reversed(self.stack):
            if not node_id:
                continue
            node = self.nodes.get(node_id)
            if node and node.kind == "run":
                node.append_run_text(data, active_segment_attrs)
                return
            if node and node.kind == "protected":
                node.text_parts.append(data)
                return


def collect_hdocx_ids(html_text: str) -> list[str]:
    scanner = HDocxHTMLScanner()
    scanner.feed(html_text)
    return scanner.ids


def collect_hdocx_nodes(html_text: str) -> dict[str, HTMLNodeSnapshot]:
    scanner = HDocxHTMLScanner()
    scanner.feed(html_text)
    return scanner.nodes
