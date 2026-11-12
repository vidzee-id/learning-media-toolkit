import io
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import imageio_ffmpeg
import streamlit as st
from pptx import Presentation
from docx import Document
from docx.shared import Inches, Pt


st.set_page_config(
    page_title="Learning Media Toolkit",
    page_icon="🎓",
    layout="wide",
)

# Friendly, lightweight styling. Uses standard Streamlit components.
st.markdown(
    """
    <style>
      .stApp {
        background: linear-gradient(180deg, #f5f8ff 0%, #ffffff 42%);
      }
      .block-container {
        max-width: 1120px;
        padding-top: 1.8rem;
        padding-bottom: 3rem;
      }
      .hero {
        background: linear-gradient(115deg, #2447a8 0%, #4677dc 58%, #6d8fe8 100%);
        border-radius: 22px;
        padding: 1.6rem 1.8rem;
        color: white;
        margin-bottom: 1.25rem;
        box-shadow: 0 10px 28px rgba(36, 71, 168, 0.15);
      }
      .hero h1 {
        color: white;
        font-size: 2rem;
        line-height: 1.2;
        margin: 0 0 .45rem 0;
      }
      .hero p {
        color: rgba(255,255,255,.92);
        margin: 0;
        font-size: 1rem;
      }
      .tool-card {
        background: white;
        border: 1px solid #e5eaf5;
        border-radius: 18px;
        padding: 1.1rem 1.25rem;
        margin: .2rem 0 1rem 0;
        box-shadow: 0 4px 16px rgba(27, 50, 100, .045);
      }
      .tool-icon {
        display: inline-flex;
        width: 46px;
        height: 46px;
        border-radius: 14px;
        align-items: center;
        justify-content: center;
        font-size: 1.5rem;
        background: #edf2ff;
        margin-bottom: .45rem;
      }
      .tool-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #1d2b53;
        margin: .1rem 0 .25rem 0;
      }
      .tool-description {
        color: #5c6780;
        font-size: .95rem;
        margin: 0;
      }
      .small-note {
        color: #667085;
        font-size: .88rem;
      }
      div[data-testid="stFileUploader"] {
        border: 1px dashed #b9c8ee;
        border-radius: 14px;
        padding: .45rem .6rem;
        background: #fbfcff;
      }
      .stButton > button, .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 600;
        min-height: 2.7rem;
      }
      div[data-testid="stMetric"] {
        background: #f7f9ff;
        border: 1px solid #e6ebfa;
        padding: .8rem;
        border-radius: 12px;
      }
      div[data-testid="stTabs"] button {
        border-radius: 10px 10px 0 0;
        font-weight: 600;
      }
      hr { border-color: #e8edf6; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <h1>🎓 Learning Media Toolkit</h1>
      <p>Your everyday toolkit for learning media production. Extract, convert and prepare content with simple, practical tools.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:12px;margin:0 0 1.35rem 0;">
      <div class="tool-card" style="margin:0;">
        <div class="tool-icon">📝</div>
        <div class="tool-title">PPT Notes Extractor</div>
        <p class="tool-description">Extract speaker notes and export them to Word or TXT.</p>
      </div>
      <div class="tool-card" style="margin:0;">
        <div class="tool-icon" style="background:#fff2ea;">🔊</div>
        <div class="tool-title">PPT Audio Extractor</div>
        <p class="tool-description">Collect slide audio in order and merge it into one MP3.</p>
      </div>
      <div class="tool-card" style="margin:0;">
        <div class="tool-icon" style="background:#eaf7f0;">🎧</div>
        <div class="tool-title">MP4 to MP3 Converter</div>
        <p class="tool-description">Extract the audio track from an MP4 video.</p>
      </div>
    </div>
    <p class="small-note" style="margin:0 0 1rem 0;">Choose a utility below to get started.</p>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------
# PowerPoint speaker notes extraction
# -------------------------------------------------------------------
def _local_name(tag):
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _extract_text_from_notes_xml(notes_xml_bytes):
    root = ET.fromstring(notes_xml_bytes)
    paragraphs = []

    for paragraph in root.iter():
        if _local_name(paragraph.tag) != "p":
            continue

        pieces = []
        for node in paragraph.iter():
            name = _local_name(node.tag)
            if name == "t" and node.text:
                pieces.append(node.text)
            elif name == "tab":
                pieces.append("\t")
            elif name == "br":
                pieces.append("\n")

        text = "".join(pieces).strip()
        if text:
            paragraphs.append(text)

    if not paragraphs:
        pieces = []
        for node in root.iter():
            if _local_name(node.tag) == "t" and node.text:
                pieces.append(node.text)
        fallback = " ".join(pieces).strip()
        if fallback:
            paragraphs.append(fallback)

    cleaned = []
    for item in paragraphs:
        item = re.sub(r"[ \t]+", " ", item).strip()
        if item and (not cleaned or item != cleaned[-1]):
            cleaned.append(item)

    return "\n".join(cleaned).strip()


def extract_ppt_notes(pptx_file):
    result = {
        "notes_text": "",
        "total_slides": 0,
        "extracted_slides": [],
        "empty_slides": [],
        "failed_slides": [],
        "slide_notes": {},
    }

    pptx_bytes = pptx_file.getvalue()
    presentation = Presentation(io.BytesIO(pptx_bytes))
    result["total_slides"] = len(presentation.slides)

    try:
        with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as archive:
            names = set(archive.namelist())
            slide_notes = {}

            for slide_number in range(1, result["total_slides"] + 1):
                slide_path = f"ppt/slides/slide{slide_number}.xml"
                rel_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"

                try:
                    if slide_path not in names:
                        result["failed_slides"].append(slide_number)
                        continue

                    notes_path = None
                    if rel_path in names:
                        rel_root = ET.fromstring(archive.read(rel_path))
                        for rel in rel_root:
                            if rel.attrib.get("Type", "").endswith("/notesSlide"):
                                target = rel.attrib.get("Target", "")
                                if target:
                                    if target.startswith("/"):
                                        notes_path = target.lstrip("/")
                                    else:
                                        notes_path = str(
                                            Path("ppt/slides") / target
                                        ).replace("\\", "/")
                                        parts = []
                                        for part in notes_path.split("/"):
                                            if part == "..":
                                                if parts:
                                                    parts.pop()
                                            elif part != ".":
                                                parts.append(part)
                                        notes_path = "/".join(parts)
                                    break

                    if not notes_path or notes_path not in names:
                        result["empty_slides"].append(slide_number)
                        continue

                    notes = _extract_text_from_notes_xml(archive.read(notes_path))
                    if notes.strip().lower() in {
                        "click to add notes",
                        "click to edit master text styles",
                    }:
                        notes = ""

                    if notes:
                        result["extracted_slides"].append(slide_number)
                        slide_notes[slide_number] = notes
                    else:
                        result["empty_slides"].append(slide_number)

                except Exception:
                    result["failed_slides"].append(slide_number)

            result["slide_notes"] = slide_notes
            output_blocks = [
                f"SLIDE {number}\n{slide_notes[number]}"
                for number in range(1, result["total_slides"] + 1)
                if number in slide_notes
            ]
            result["notes_text"] = "\n\n".join(output_blocks)

    except zipfile.BadZipFile as exc:
        raise ValueError(
            "This file does not appear to be a valid, unencrypted .pptx file."
        ) from exc
    except Exception as exc:
        raise ValueError(f"Could not read the PowerPoint file: {exc}") from exc

    return result


def create_notes_word_document(result, source_filename):
    """Build a formatted Word document containing extracted notes by slide."""
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    title = document.add_heading("PowerPoint Speaker Notes", level=0)
    title.style.font.name = "Aptos Display"
    subtitle = document.add_paragraph()
    subtitle.add_run("Source presentation: ").bold = True
    subtitle.add_run(source_filename)
    subtitle.add_run("\nSlides in presentation: ").bold = True
    subtitle.add_run(str(result["total_slides"]))
    subtitle.add_run("\nSlides with notes extracted: ").bold = True
    subtitle.add_run(str(len(result["extracted_slides"])))

    document.add_paragraph(
        "Review the extracted notes before use. Slides without readable notes may be blank or may contain content that needs manual checking."
    )

    slide_notes = result.get("slide_notes", {})
    for slide_number in sorted(slide_notes):
        document.add_heading(f"Slide {slide_number}", level=1)
        for paragraph_text in slide_notes[slide_number].splitlines():
            if paragraph_text.strip():
                document.add_paragraph(paragraph_text.strip())

    if result.get("empty_slides") or result.get("failed_slides"):
        document.add_heading("Slides to Review", level=1)
        if result.get("empty_slides"):
            document.add_paragraph(
                "No readable notes found on slide(s): "
                + ", ".join(map(str, result["empty_slides"]))
            )
        if result.get("failed_slides"):
            document.add_paragraph(
                "Could not read slide(s): "
                + ", ".join(map(str, result["failed_slides"]))
            )

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output.getvalue()


# -------------------------------------------------------------------
# MP4 to MP3 conversion
# -------------------------------------------------------------------
def convert_mp4_to_mp3(uploaded_video):
    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    video_bytes = uploaded_video.getvalue()

    with tempfile.TemporaryDirectory() as temp_dir:
        input_path = Path(temp_dir) / "input.mp4"
        output_path = Path(temp_dir) / "output.mp3"
        input_path.write_bytes(video_bytes)

        command = [
            ffmpeg_exe,
            "-hide_banner",
            "-loglevel", "error",
            "-y",
            "-i", str(input_path),
            "-vn",
            "-codec:a", "libmp3lame",
            "-q:a", "2",
            str(output_path),
        ]

        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=600,
        )

        if completed.returncode != 0:
            detail = (
                completed.stderr or "FFmpeg could not convert this file."
            ).strip()
            raise ValueError(detail[-1800:])

        if not output_path.exists() or output_path.stat().st_size == 0:
            raise ValueError(
                "No audio was produced. The MP4 may not contain an audio track."
            )

        return output_path.read_bytes()


# -------------------------------------------------------------------
# PowerPoint slide audio extraction and merging
# -------------------------------------------------------------------
AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".aac", ".wma", ".oga", ".ogg", ".opus", ".mid", ".midi"}


def _normalise_zip_target(base_folder, target):
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base_folder + "/" + target).split("/")
    cleaned = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if cleaned:
                cleaned.pop()
        else:
            cleaned.append(part)
    return "/".join(cleaned)


def extract_audio_references(pptx_bytes):
    """Return embedded audio media references in presentation slide order."""
    presentation = Presentation(io.BytesIO(pptx_bytes))
    found = []
    slides_without_audio = []

    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as archive:
        archive_names = set(archive.namelist())
        for slide_number, slide in enumerate(presentation.slides, start=1):
            slide_path = str(slide.part.partname).lstrip("/")
            slide_xml = ET.fromstring(archive.read(slide_path))
            rels_path = str(Path(slide_path).parent / "_rels" / (Path(slide_path).name + ".rels"))
            rel_map = {}
            if rels_path in archive_names:
                rel_root = ET.fromstring(archive.read(rels_path))
                for rel in rel_root:
                    rid = rel.attrib.get("Id")
                    target = rel.attrib.get("Target", "")
                    rel_type = rel.attrib.get("Type", "").lower()
                    if rid and target:
                        rel_map[rid] = {
                            "path": _normalise_zip_target(str(Path(slide_path).parent), target),
                            "type": rel_type,
                            "mode": rel.attrib.get("TargetMode", ""),
                        }

            slide_clips = []
            seen_rids = set()
            seen_media_paths = set()
            # PowerPoint may expose the same embedded audio through multiple
            # relationship IDs (for example, an audio object and its media
            # reference). Deduplicate by the actual package media path within
            # each slide, while still allowing the same file on another slide.
            # Reading attributes in document order preserves slide clip order.
            for element in slide_xml.iter():
                for attr_name, rid in element.attrib.items():
                    if not attr_name.endswith("}embed") and not attr_name.endswith("}link"):
                        continue
                    if rid in seen_rids or rid not in rel_map:
                        continue
                    seen_rids.add(rid)
                    rel = rel_map[rid]
                    media_path = rel["path"]
                    suffix = Path(media_path).suffix.lower()
                    is_audio = suffix in AUDIO_EXTENSIONS or "/audio" in rel["type"]
                    if (
                        is_audio
                        and rel["mode"] != "External"
                        and media_path in archive_names
                        and media_path not in seen_media_paths
                    ):
                        seen_media_paths.add(media_path)
                        slide_clips.append({
                            "slide": slide_number,
                            "relationship_id": rid,
                            "media_path": media_path,
                            "data": archive.read(media_path),
                        })

            if slide_clips:
                found.extend(slide_clips)
            else:
                slides_without_audio.append(slide_number)

    return found, slides_without_audio, len(presentation.slides)


def create_merged_ppt_audio(uploaded_pptx):
    """Extract slide audio, normalize clips, merge in slide order, and return MP3."""
    pptx_bytes = uploaded_pptx.getvalue()
    clips, slides_without_audio, total_slides = extract_audio_references(pptx_bytes)
    if not clips:
        raise ValueError(
            "No embedded audio clips were found. Check that audio is inserted into the slides, "
            "rather than linked from an external file."
        )

    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        wav_paths = []
        clip_report = []

        for index, clip in enumerate(clips, start=1):
            suffix = Path(clip["media_path"]).suffix.lower() or ".bin"
            input_path = temp_path / f"clip_{index:04d}{suffix}"
            wav_path = temp_path / f"clip_{index:04d}.wav"
            input_path.write_bytes(clip["data"])

            command = [
                ffmpeg_exe, "-hide_banner", "-loglevel", "error", "-y",
                "-i", str(input_path), "-vn", "-ac", "2", "-ar", "44100",
                "-c:a", "pcm_s16le", str(wav_path),
            ]
            completed = subprocess.run(command, capture_output=True, text=True, timeout=600)
            if completed.returncode != 0 or not wav_path.exists() or wav_path.stat().st_size == 0:
                detail = (completed.stderr or "FFmpeg could not read this audio clip.").strip()
                raise ValueError(f"Could not process audio on slide {clip['slide']} ({clip['media_path']}): {detail[-900:]}")

            wav_paths.append(wav_path)
            clip_report.append({
                "Slide": clip["slide"],
                "Clip": index,
                "Embedded file": Path(clip["media_path"]).name,
            })

        concat_file = temp_path / "audio_list.txt"
        concat_file.write_text(
            "".join(f"file '{path}'\n" for path in wav_paths),
            encoding="utf-8",
        )
        output_path = temp_path / "merged_audio.mp3"
        merge_command = [
            ffmpeg_exe, "-hide_banner", "-loglevel", "error", "-y",
            "-f", "concat", "-safe", "0", "-i", str(concat_file),
            "-c:a", "libmp3lame", "-q:a", "2", str(output_path),
        ]
        completed = subprocess.run(merge_command, capture_output=True, text=True, timeout=900)
        if completed.returncode != 0 or not output_path.exists() or output_path.stat().st_size == 0:
            detail = (completed.stderr or "FFmpeg could not merge the slide audio.").strip()
            raise ValueError(detail[-1500:])

        return output_path.read_bytes(), clip_report, slides_without_audio, total_slides


# -------------------------------------------------------------------
# Main tools
# -------------------------------------------------------------------
ppt_tab, ppt_audio_tab, audio_tab = st.tabs(
    ["📝  PPT Notes Extractor", "🔊  PPT Audio Extractor", "🎧  MP4 to MP3"]
)

with ppt_tab:
    st.markdown(
        """
        <div class="tool-card">
          <div class="tool-icon">📝</div>
          <div class="tool-title">PPT Notes Extractor</div>
          <p class="tool-description">Pull speaker notes into a clean text file, with a quick report of slides that may need checking.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_pptx = st.file_uploader(
        "Choose a PowerPoint file",
        type=["pptx"],
        help="Upload a .pptx file. Password-protected files are not supported.",
        key="pptx_upload",
    )

    if uploaded_pptx is not None:
        st.caption(f"📄 Selected file: {uploaded_pptx.name}")

        if st.button("✨ Extract speaker notes", key="extract_notes_button", type="primary"):
            try:
                with st.spinner("Reading slides and collecting notes…"):
                    result = extract_ppt_notes(uploaded_pptx)

                st.success("Done! Your PowerPoint has been processed.")

                col1, col2, col3 = st.columns(3)
                col1.metric("📑 Total slides", result["total_slides"])
                col2.metric("✅ Notes extracted", len(result["extracted_slides"]))
                col3.metric(
                    "🔎 Slides to review",
                    len(result["empty_slides"]) + len(result["failed_slides"]),
                )

                if result["empty_slides"]:
                    slide_list = ", ".join(map(str, result["empty_slides"]))
                    st.warning(
                        "No readable notes were extracted from slide(s): "
                        f"**{slide_list}**. They may have no notes, or may contain "
                        "equations or other content that needs manual review."
                    )

                if result["failed_slides"]:
                    slide_list = ", ".join(map(str, result["failed_slides"]))
                    st.error(
                        f"Could not read slide(s): **{slide_list}**. "
                        "Try opening and resaving the presentation in PowerPoint."
                    )

                if result["notes_text"].strip():
                    st.markdown("#### 📋 Extracted notes")
                    st.text_area(
                        "Review the notes below before downloading",
                        value=result["notes_text"],
                        height=360,
                        label_visibility="collapsed",
                    )
                    safe_name = Path(uploaded_pptx.name).stem
                    download_col1, download_col2 = st.columns(2)
                    with download_col1:
                        st.download_button(
                            "⬇️ Download notes as TXT",
                            data=result["notes_text"].encode("utf-8"),
                            file_name=f"{safe_name}_speaker_notes.txt",
                            mime="text/plain",
                            type="primary",
                        )
                    with download_col2:
                        try:
                            word_bytes = create_notes_word_document(
                                result, uploaded_pptx.name
                            )
                            st.download_button(
                                "📝 Export to Word (.docx)",
                                data=word_bytes,
                                file_name=f"{safe_name}_speaker_notes.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            )
                        except Exception as exc:
                            st.error("Could not create the Word document.")
                            with st.expander("Show Word export details"):
                                st.code(str(exc))
                else:
                    st.info(
                        "No readable speaker notes were found. Please check the "
                        "slide review information above and verify the notes in PowerPoint."
                    )

            except Exception as exc:
                st.error(
                    "We couldn't process this PowerPoint file. Please check that it "
                    "is a valid, unencrypted PPTX file."
                )
                with st.expander("Show technical details"):
                    st.code(str(exc))

    else:
        st.markdown(
            '<p class="small-note">💡 Tip: speaker notes are different from the text visible on the slides.</p>',
            unsafe_allow_html=True,
        )


with ppt_audio_tab:
    st.markdown(
        """
        <div class="tool-card">
          <div class="tool-icon">🔊</div>
          <div class="tool-title">PPT Audio Extractor & Merger</div>
          <p class="tool-description">Extract embedded audio from every slide, arrange clips in slide order, and combine them into one MP3.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_audio_pptx = st.file_uploader(
        "Choose a PowerPoint file",
        type=["pptx"],
        help="Upload a .pptx presentation containing audio clips embedded in its slides.",
        key="ppt_audio_upload",
    )

    if uploaded_audio_pptx is not None:
        st.caption(f"📄 Selected file: {uploaded_audio_pptx.name} · {uploaded_audio_pptx.size / (1024 * 1024):.1f} MB")
        st.info("Audio clips are combined in slide order. If a slide contains multiple audio objects, their order in the slide's audio references is used. External linked audio files cannot be retrieved.")

        if st.button("✨ Extract and merge slide audio", key="merge_ppt_audio_button", type="primary"):
            try:
                with st.spinner("Finding slide audio, converting clips and creating your MP3…"):
                    mp3_bytes, clip_report, slides_without_audio, total_slides = create_merged_ppt_audio(uploaded_audio_pptx)

                st.success("Your PowerPoint audio has been merged successfully!")
                col1, col2, col3 = st.columns(3)
                col1.metric("📑 Total slides", total_slides)
                col2.metric("🔊 Audio clips", len(clip_report))
                col3.metric("🔎 Slides without audio", len(slides_without_audio))

                if slides_without_audio:
                    st.warning("No embedded audio was found on slide(s): " + ", ".join(map(str, slides_without_audio)))

                with st.expander("View extracted audio clip list"):
                    st.dataframe(clip_report, use_container_width=True, hide_index=True)

                safe_name = Path(uploaded_audio_pptx.name).stem
                st.audio(mp3_bytes, format="audio/mpeg")
                st.download_button(
                    "⬇️ Download complete audio as MP3",
                    data=mp3_bytes,
                    file_name=f"{safe_name}_complete_audio.mp3",
                    mime="audio/mpeg",
                    type="primary",
                )
            except subprocess.TimeoutExpired:
                st.error("Audio processing took too long. Try a smaller presentation or fewer/lower-size audio clips.")
            except Exception as exc:
                st.error("We couldn't extract or merge the slide audio.")
                with st.expander("Show technical details"):
                    st.code(str(exc))
    else:
        st.markdown(
            '<p class="small-note">💡 Tip: use a .pptx file with audio inserted into the slides. Audio linked from an external location will not be included.</p>',
            unsafe_allow_html=True,
        )


with audio_tab:
    st.markdown(
        """
        <div class="tool-card">
          <div class="tool-icon">🎧</div>
          <div class="tool-title">MP4 to MP3 Converter</div>
          <p class="tool-description">Extract the audio track from a video and save it as an MP3 file.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_mp4 = st.file_uploader(
        "Choose an MP4 video",
        type=["mp4"],
        help="Upload an MP4 file that contains an audio track.",
        key="mp4_upload",
    )

    if uploaded_mp4 is not None:
        size_mb = uploaded_mp4.size / (1024 * 1024)
        st.caption(f"🎬 Selected file: {uploaded_mp4.name}  ·  {size_mb:.1f} MB")

        if st.button("🎵 Convert to MP3", key="convert_mp3_button", type="primary"):
            try:
                with st.spinner("Extracting audio… larger videos may take a few minutes."):
                    mp3_bytes = convert_mp4_to_mp3(uploaded_mp4)

                safe_name = Path(uploaded_mp4.name).stem
                st.success("Your MP3 is ready!")
                st.download_button(
                    "⬇️ Download MP3",
                    data=mp3_bytes,
                    file_name=f"{safe_name}.mp3",
                    mime="audio/mpeg",
                    type="primary",
                )
            except subprocess.TimeoutExpired:
                st.error("The conversion took too long. Try a shorter video.")
            except Exception as exc:
                st.error("We couldn't convert this MP4 file.")
                with st.expander("Show technical details"):
                    st.code(str(exc))

    else:
        st.markdown(
            '<p class="small-note">💡 Tip: the MP3 will usually be much smaller than the original video.</p>',
            unsafe_allow_html=True,
        )

st.markdown("---")
st.markdown(
    '<p class="small-note" style="text-align:center;">Learning Media Toolkit · Practical tools for instructional design and learning media production</p>',
    unsafe_allow_html=True,
)
