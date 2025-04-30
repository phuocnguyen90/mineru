# app.py
import streamlit as st
import subprocess
import os
import shutil
from pathlib import Path
import base64
from docx import Document
from docx.shared import Inches
import uuid # For unique job IDs
import hashlib # For hashing file content
import json # For status file

st.set_page_config(layout="wide")
# --- Configuration ---
# Base directories (Colab ephemeral storage is the fallback)
TEMP_BASE_DIR = Path("/content/temp_processing")
DEFAULT_UPLOAD_DIR = TEMP_BASE_DIR / "uploads"
DEFAULT_OUTPUT_DIR = TEMP_BASE_DIR / "output"
STATUS_FILENAME = "processing_status.json"

# Google Drive path (set by Colab notebook via environment variable)
DRIVE_APP_BASE_PATH_STR = os.environ.get('DRIVE_APP_BASE_PATH', "")
DRIVE_AVAILABLE = bool(DRIVE_APP_BASE_PATH_STR) and Path(DRIVE_APP_BASE_PATH_STR).is_dir()

if DRIVE_AVAILABLE:
    APP_BASE_DIR = Path(DRIVE_APP_BASE_PATH_STR)
    st.sidebar.success(f"✅ Google Drive Connected: {APP_BASE_DIR}")
else:
    APP_BASE_DIR = TEMP_BASE_DIR
    st.sidebar.warning("⚠️ Google Drive not connected. Files will be temporary.")

UPLOAD_DIR = APP_BASE_DIR / "uploads"
OUTPUT_DIR = APP_BASE_DIR / "output"

# Create directories
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Language mapping for magic-pdf
LANGUAGES = {
    "Auto Detect": None,
    "English": "en",
    "Chinese (Simplified)": "zh",
    "Vietnamese": "vi",
    # Add more languages supported by magic-pdf here if needed
    "Japanese": "ja",
    "Korean": "ko",
}

# --- Helper Functions ---

def get_file_hash(file_path: Path) -> str:
    """Calculates the SHA256 hash of a file."""
    hasher = hashlib.sha256()
    with open(file_path, 'rb') as file:
        while chunk := file.read(4096): # Process in chunks
            hasher.update(chunk)
    return hasher.hexdigest()

def save_uploaded_file(uploaded_file) -> tuple[Path | None, str | None]:
    """Saves uploaded file, calculates hash, returns path and hash."""
    try:
        # Use a unique temp name first to calculate hash before permanent save
        temp_path = DEFAULT_UPLOAD_DIR / f"temp_{uuid.uuid4()}"
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        file_hash = get_file_hash(temp_path)
        # Use hash for the final filename to somewhat handle duplicates/resuming
        # Append original extension
        original_ext = Path(uploaded_file.name).suffix
        permanent_path = UPLOAD_DIR / f"{file_hash}{original_ext}"

        # Move temp file to permanent location
        shutil.move(str(temp_path), str(permanent_path))
        st.info(f"Saved uploaded file to: {permanent_path}")
        return permanent_path, file_hash
    except Exception as e:
        st.error(f"Error saving uploaded file: {e}")
        if 'temp_path' in locals() and temp_path.exists():
            temp_path.unlink() # Clean up temp file
        return None, None
    finally:
        # Ensure temp file is removed if it still exists for some reason
         if 'temp_path' in locals() and temp_path.exists():
            temp_path.unlink()


def get_job_dirs(file_hash: str) -> tuple[Path, Path, Path]:
    """Gets the specific output and status directories for a job hash."""
    job_output_base = OUTPUT_DIR / file_hash
    # e.g., /content/drive/MyDrive/App/output/HASH/HASH_origin/auto/
    results_dir = job_output_base / f"{file_hash}_origin" / "auto"
    status_file_path = job_output_base / STATUS_FILENAME
    job_output_base.mkdir(parents=True, exist_ok=True)
    return job_output_base, results_dir, status_file_path

def read_status(status_file_path: Path) -> dict:
    """Reads the status JSON file."""
    if status_file_path.is_file():
        try:
            with open(status_file_path, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            st.warning(f"Status file {status_file_path} is corrupted. Starting fresh.")
            return {}
        except Exception as e:
            st.error(f"Error reading status file: {e}")
            return {}
    return {}

def write_status(status_file_path: Path, status_data: dict):
    """Writes the status JSON file."""
    try:
        with open(status_file_path, 'w') as f:
            json.dump(status_data, f, indent=2)
    except Exception as e:
        st.error(f"Error writing status file: {e}")


def run_magic_pdf(input_pdf_path: Path, job_output_base: Path, lang_code: str | None):
    """Runs the magic-pdf command and returns the results directory path."""
    command = [
        "magic-pdf",
        "-p", str(input_pdf_path),
        "-o", str(job_output_base), # Output to the specific job directory
        "-m", "auto"
    ]
    if lang_code:
        command.extend(["-l", lang_code])

    st.info(f"Running command: {' '.join(command)}")
    full_stdout = []
    full_stderr = []
    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8' # Explicitly set encoding
        )

        # Stream output to UI (optional, can be verbose)
        stdout_container = st.expander("Magic-PDF Output (stdout)", expanded=False)
        stderr_container = st.expander("Magic-PDF Errors (stderr)", expanded=True)
        with stdout_container:
            stdout_area = st.empty()
        with stderr_container:
            stderr_area = st.empty()

        # Read output line by line
        while True:
            stdout_line = process.stdout.readline()
            stderr_line = process.stderr.readline()

            if stdout_line:
                full_stdout.append(stdout_line.strip())
                stdout_area.text("\n".join(full_stdout[-20:])) # Show last 20 lines
            if stderr_line:
                full_stderr.append(stderr_line.strip())
                stderr_area.text("\n".join(full_stderr)) # Show all stderr

            # Check if process finished
            if stdout_line == '' and stderr_line == '' and process.poll() is not None:
                break

        returncode = process.returncode

        # Final display of full logs if needed
        # stdout_area.text("\n".join(full_stdout))
        # stderr_area.text("\n".join(full_stderr))

        if returncode != 0:
            st.error(f"magic-pdf command failed with return code {returncode}")
            # Combine stderr for the status file
            error_message = "\n".join(full_stderr) if full_stderr else f"Exited with code {returncode}"
            return None, error_message

        # Construct the expected path to the actual results directory
        results_dir = job_output_base / f"{input_pdf_path.stem}_origin" / "auto"
        st.success(f"Processing likely complete. Results expected in: {results_dir}")

        if not results_dir.is_dir():
             st.error(f"Expected results directory not found: {results_dir}")
             return None, f"Results directory missing after processing: {results_dir}"

        return results_dir, None # Success

    except FileNotFoundError:
        st.error("Error: 'magic-pdf' command not found. Is it installed correctly in the environment?")
        return None, "magic-pdf command not found"
    except Exception as e:
        st.error(f"An exception occurred while running magic-pdf: {e}")
        return None, str(e)

# ... (get_download_link and create_word_doc functions remain the same as before) ...

def get_download_link(file_path: Path, link_text: str, filename: str):
    """Generates a download link for a file."""
    if not file_path.is_file():
        return f"*{filename} not found*"
    try:
        with open(file_path, "rb") as f:
            bytes_data = f.read()
        b64 = base64.b64encode(bytes_data).decode()
        return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{link_text}</a>'
    except Exception as e:
        return f"*Error creating download link for {filename}: {e}*"

def create_word_doc(md_content: str, output_docx_path: Path):
    """Creates a basic Word document from Markdown text content."""
    try:
        document = Document()
        # Basic paragraph splitting (improve parsing as needed)
        for paragraph_text in md_content.split('\n'):
            if paragraph_text.strip(): # Avoid empty paragraphs
                 document.add_paragraph(paragraph_text)
        document.save(output_docx_path)
        return True
    except Exception as e:
        st.error(f"Error creating Word document: {e}")
        return False


# --- Streamlit App UI ---

st.title("📄 MinerU / Magic-PDF Processor")
st.markdown("Upload a PDF file, select the language, and process it using `magic-pdf`.")

# --- Session State Initialization ---
if 'job_id' not in st.session_state:
    st.session_state.job_id = None
if 'file_hash' not in st.session_state:
    st.session_state.file_hash = None
if 'input_pdf_path' not in st.session_state:
    st.session_state.input_pdf_path = None
if 'results_dir' not in st.session_state:
    st.session_state.results_dir = None
if 'processing_done' not in st.session_state:
    st.session_state.processing_done = False
if 'error_message' not in st.session_state:
    st.session_state.error_message = None

# --- UI Elements ---
uploaded_file = st.file_uploader("1. Choose a PDF file", type="pdf", key="pdf_uploader")
selected_language_name = st.selectbox(
    "2. Select Language (or Auto Detect)",
    options=list(LANGUAGES.keys()),
    key="lang_select"
)
lang_code = LANGUAGES[selected_language_name]

process_button = st.button("✨ Process PDF", key="process_button", disabled=(uploaded_file is None))

# --- Processing Logic ---
if process_button and uploaded_file is not None:
    # Reset previous results
    st.session_state.job_id = str(uuid.uuid4()) # New job ID for this run
    st.session_state.results_dir = None
    st.session_state.processing_done = False
    st.session_state.error_message = None

    # Save file and get hash
    saved_path, file_hash = save_uploaded_file(uploaded_file)

    if saved_path and file_hash:
        st.session_state.input_pdf_path = str(saved_path)
        st.session_state.file_hash = file_hash

        job_output_base, expected_results_dir, status_file_path = get_job_dirs(file_hash)

        # --- Resume Logic (Basic Example) ---
        st.info(f"Checking status at: {status_file_path}")
        status_data = read_status(status_file_path)
        if status_data.get("status") == "completed" and Path(status_data.get("results_dir", "")).is_dir():
             st.success("Found completed results from a previous run for this file hash.")
             st.session_state.results_dir = Path(status_data["results_dir"])
             st.session_state.processing_done = True
             st.session_state.error_message = None
             # Force rerun to display results
             st.experimental_rerun()
        else:
            st.info("Starting new processing job...")
            status_data = {
                "job_id": st.session_state.job_id,
                "file_hash": file_hash,
                "original_filename": uploaded_file.name,
                "language": lang_code if lang_code else "auto",
                "status": "started",
                "results_dir": None,
                "error": None
            }
            write_status(status_file_path, status_data)

            with st.spinner(f"Processing PDF with magic-pdf (lang={lang_code or 'auto'})... This can take several minutes."):
                results_dir_path, error_msg = run_magic_pdf(saved_path, job_output_base, lang_code)

            if results_dir_path and not error_msg:
                st.session_state.results_dir = results_dir_path
                st.session_state.processing_done = True
                st.session_state.error_message = None
                status_data["status"] = "completed"
                status_data["results_dir"] = str(results_dir_path)
                write_status(status_file_path, status_data)
                st.success("Processing finished successfully!")
                 # Force rerun to display results section
                st.experimental_rerun()
            else:
                st.session_state.error_message = error_msg or "Unknown processing error."
                status_data["status"] = "error"
                status_data["error"] = st.session_state.error_message
                write_status(status_file_path, status_data)
                st.error(f"Processing failed: {st.session_state.error_message}")
                 # Force rerun to display error section
                st.experimental_rerun()

# --- Display Results or Errors (runs after button press or rerun) ---
if st.session_state.processing_done and st.session_state.results_dir:
    st.markdown("---")
    st.header("📊 Processing Results")
    results_dir = Path(st.session_state.results_dir) # Ensure it's a Path object
    input_pdf_stem = Path(st.session_state.input_pdf_path).stem

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Extracted Markdown")
        md_file_path = results_dir / "result.md"
        if md_file_path.is_file():
            try:
                md_content = md_file_path.read_text(encoding='utf-8')
                st.markdown(md_content, unsafe_allow_html=True) # Display Markdown

                # Prepare Word doc
                temp_docx_path = results_dir / f"{input_pdf_stem}_output.docx"
                docx_link = ""
                if create_word_doc(md_content, temp_docx_path):
                     docx_link = get_download_link(temp_docx_path, "Download Word (.docx)", f"{input_pdf_stem}.docx")

                md_link = get_download_link(md_file_path, "Download Markdown (.md)", f"{input_pdf_stem}.md")

                st.markdown("---") # Separator
                st.markdown(f"{md_link} | {docx_link}", unsafe_allow_html=True)


            except Exception as e:
                st.error(f"Error reading or displaying Markdown file: {e}")
        else:
            st.warning(f"Could not find result.md in {results_dir}")

    with col2:
        st.subheader("🖼️ Page Analysis (Images)")
        image_files = sorted(list(results_dir.glob("*.png"))) # Look for PNGs
        if image_files:
            st.info(f"Found {len(image_files)} preview image(s).")
            # Use tabs for better organization if many images
            image_tabs = st.tabs([f"Page {i+1}" for i in range(len(image_files))])
            for i, img_path in enumerate(image_files):
                 with image_tabs[i]:
                    try:
                        st.image(str(img_path), caption=img_path.name, use_column_width=True)
                    except Exception as e:
                        st.error(f"Could not display image {img_path.name}: {e}")
        else:
            st.info("No preview images found in the output directory.")

    # Display path to full output
    st.markdown(f"--- \n*Full output files are located in:* `{results_dir}`")

elif st.session_state.error_message:
    st.markdown("---")
    st.error(f"Processing failed with error: {st.session_state.error_message}")
    st.info("Check the stderr output in the expander during the run for more details.")

# --- Footer/Info ---
st.sidebar.markdown("---")
st.sidebar.info(f"Using {'Google Drive' if DRIVE_AVAILABLE else 'Temporary Colab Storage'}")
st.sidebar.info(f"Base Path: {APP_BASE_DIR}")