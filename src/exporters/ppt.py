"""PowerPoint exporter using python-pptx."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from src.schemas.lesson import LessonPackage
from templates.ppt_template import apply_content_slide, apply_title_slide


class PptxExporter:
    """Export a LessonPackage to a .pptx file."""

    def __init__(self) -> None:
        self._presentation: Presentation | None = None

    def export(self, package: LessonPackage, output_path: Path | str) -> Path:
        out = Path(output_path)
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)

        blank_layout = prs.slide_layouts[6]  # blank
        ref_lookup = {r.citation_id: r for r in package.references}

        for idx, slide in enumerate(package.slides):
            s = prs.slides.add_slide(blank_layout)
            footer = self._build_footer(slide.citation_ids, ref_lookup)

            if idx == 0 or slide.slide_type == "title":
                apply_title_slide(
                    s,
                    title=slide.title,
                    subtitle=package.metadata.audience,
                )
            else:
                apply_content_slide(
                    s,
                    title=slide.title,
                    bullets=slide.key_points,
                    footer=footer,
                )

            # Add speaker notes if python-pptx supports them.
            if slide.speaker_notes:
                try:
                    notes_slide = s.notes_slide
                    notes_frame = notes_slide.notes_text_frame
                    notes_frame.text = slide.speaker_notes
                except Exception:
                    pass  # Notes are also exported to speaker_notes.md.

        out.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out))
        return out

    @staticmethod
    def _build_footer(
        citation_ids: list[str],
        ref_lookup: dict[str, "CitationRecord"],
    ) -> str:
        if not citation_ids:
            return ""
        short_refs = []
        for cid in citation_ids:
            ref = ref_lookup.get(cid)
            if ref:
                if ref.reference_type.value == "pubmed_retrieved" and ref.pmid:
                    short_refs.append(f"PMID: {ref.pmid}")
                else:
                    year = f" {ref.year}" if ref.year else ""
                    short_refs.append(f"{cid}{year}")
            else:
                short_refs.append(cid)
        return "Sources: " + "; ".join(short_refs)
