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
import time # Needed for sleep
import logging # For better logging

st.set_page_config(layout="wide")
# --- Configuration ---
# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

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
    # Ensure the base app directory exists on Drive
    try:
        APP_BASE_DIR.mkdir(parents=True, exist_ok=True)
        logging.info(f"Using Google Drive base path: {APP_BASE_DIR}")
    except Exception as e:
        logging.error(f"Failed to create or access Google Drive path {APP_BASE_DIR}: {e}. Falling back to temp storage.")
        DRIVE_AVAILABLE = False # Disable drive usage if creation fails
        APP_BASE_DIR = TEMP_BASE_DIR
else:
    APP_BASE_DIR = TEMP_BASE_DIR
    logging.warning("Google Drive not connected or path invalid. Using temporary Colab storage.")

UPLOAD_DIR = APP_BASE_DIR / "uploads"
OUTPUT_DIR = APP_BASE_DIR / "output"

# Create directories if they don't exist, with error handling
try:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logging.info(f"Upload directory: {UPLOAD_DIR}")
    logging.info(f"Output directory: {OUTPUT_DIR}")
except Exception as e:
    st.error(f"Fatal Error: Could not create necessary directories ({UPLOAD_DIR}, {OUTPUT_DIR}): {e}")
    st.stop() # Stop the app if we can't create essential folders


# Language mapping for magic-pdf
LANGUAGES = {
    "Auto Detect": None,
    "English": "en",
    "Chinese (Simplified)": "zh",
    "Vietnamese": "vi",
    "Japanese": "ja",
    "Korean": "ko",
}

# --- Helper Functions ---

def get_file_hash(file_path: Path) -> str:
    """Calculates the SHA256 hash of a file."""
    hasher = hashlib.sha256()
    BLOCK_SIZE = 65536 # Read in 64k chunks
    try:
        with open(file_path, 'rb') as file:
            buf = file.read(BLOCK_SIZE)
            while len(buf) > 0:
                hasher.update(buf)
                buf = file.read(BLOCK_SIZE)
        return hasher.hexdigest()
    except Exception as e:
        logging.error(f"Error calculating hash for {file_path}: {e}")
        return f"error_{uuid.uuid4()}" # Return a unique error hash

def save_uploaded_file(uploaded_file) -> tuple[Path | None, str | None, str | None]: # Return path, hash, original_stem
    """Saves uploaded file, calculates hash, returns path, hash, and original stem."""
    if not uploaded_file:
        return None, None, None
    try:
        original_filename = uploaded_file.name
        original_stem = Path(original_filename).stem
        original_ext = Path(original_filename).suffix
        logging.info(f"Processing uploaded file: {original_filename}")

        # Use a unique temp name first to calculate hash
        # Ensure temp dir exists
        DEFAULT_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        temp_path = DEFAULT_UPLOAD_DIR / f"temp_{uuid.uuid4()}{original_ext}"
        logging.info(f"Saving temporary file to: {temp_path}")
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        logging.info(f"Calculating hash for {temp_path}")
        file_hash = get_file_hash(temp_path)
        logging.info(f"File hash: {file_hash}")

        # Use hash + original stem for the final filename for uniqueness but easier debugging
        permanent_filename = f"{file_hash}_{original_stem}{original_ext}"
        permanent_path = UPLOAD_DIR / permanent_filename

        logging.info(f"Moving temp file to permanent location: {permanent_path}")
        # Ensure target dir exists
        permanent_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(temp_path), str(permanent_path))

        logging.info(f"Successfully saved uploaded file to: {permanent_path}")
        return permanent_path, file_hash, original_stem

    except Exception as e:
        logging.error(f"Error saving uploaded file: {e}", exc_info=True)
        st.error(f"Error saving uploaded file: {e}")
        if 'temp_path' in locals() and temp_path.exists():
            try:
                temp_path.unlink()
                logging.info(f"Cleaned up temporary file: {temp_path}")
            except Exception as unlink_e:
                logging.error(f"Error cleaning up temp file {temp_path}: {unlink_e}")
        return None, None, None
    finally:
         # Defensive cleanup
         if 'temp_path' in locals() and temp_path.exists():
             try:
                 temp_path.unlink()
             except Exception:
                 pass

# **** Definition Added ****
def read_status(status_file_path: Path) -> dict:
    """Reads the status JSON file."""
    if status_file_path.is_file():
        try:
            logging.info(f"Reading status file: {status_file_path}")
            with open(status_file_path, 'r', encoding='utf-8') as f:
                status_data = json.load(f)
            logging.info(f"Status read: {status_data.get('status', 'N/A')}")
            return status_data
        except json.JSONDecodeError:
            logging.warning(f"Status file {status_file_path} is corrupted or empty. Treating as new job.")
            return {} # Return empty dict if corrupted
        except Exception as e:
            logging.error(f"Error reading status file {status_file_path}: {e}")
            st.warning(f"Could not read status file: {e}. Assuming new job.")
            return {} # Return empty dict on other errors
    else:
        logging.info(f"Status file not found: {status_file_path}")
        return {} # Return empty dict if file doesn't exist

# **** Definition Added ****
def write_status(status_file_path: Path, status_data: dict):
    """Writes the status JSON file."""
    try:
        logging.info(f"Writing status '{status_data.get('status')}' to: {status_file_path}")
        status_file_path.parent.mkdir(parents=True, exist_ok=True) # Ensure directory exists
        with open(status_file_path, 'w', encoding='utf-8') as f:
            json.dump(status_data, f, indent=2)
        logging.info(f"Status written successfully.")
    except Exception as e:
        logging.error(f"Error writing status file {status_file_path}: {e}", exc_info=True)
        st.error(f"Critical Error: Could not write job status to {status_file_path}")

def get_job_dirs(file_hash: str) -> tuple[Path, Path]:
    """Gets the specific output base dir and status file path for a job hash."""
    job_output_base = OUTPUT_DIR / file_hash # Base dir for THIS specific file hash
    status_file_path = job_output_base / STATUS_FILENAME
    try:
        job_output_base.mkdir(parents=True, exist_ok=True)
    except Exception as e:
         logging.error(f"Failed to create job output directory {job_output_base}: {e}")
         st.error(f"Error creating job directory: {e}. Cannot proceed.")
         st.stop() # Stop if we can't create job dir
    return job_output_base, status_file_path


def run_magic_pdf(input_pdf_path: Path, original_stem: str, job_output_base: Path, lang_code: str | None):
    """Runs the magic-pdf command and returns the actual results directory path."""
    command = [
        "magic-pdf",
        "-p", str(input_pdf_path),
        "-o", str(job_output_base), # Output to the specific job *base* directory
        "-m", "auto"
    ]
    if lang_code:
        command.extend(["-l", lang_code])

    st.info(f"Running command: {' '.join(command)}")
    logging.info(f"Executing magic-pdf command: {' '.join(command)}")
    full_stdout = []
    full_stderr = []

    # Define the final results dir path based on the ORIGINAL filename stem
    results_dir = job_output_base / f"{original_stem}_origin" / "auto"
    logging.info(f"Expected results directory: {results_dir}")

    try:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8', # Be explicit
            errors='replace', # Handle potential encoding errors in output
            env=os.environ.copy() # Ensure PATH is inherited
        )

        log_col1, log_col2 = st.columns(2)
        with log_col1:
             stdout_expander = st.expander("Magic-PDF Output (stdout)", expanded=False)
             stdout_area = stdout_expander.empty()
        with log_col2:
             stderr_expander = st.expander("Magic-PDF Errors (stderr)", expanded=True)
             stderr_area = stderr_expander.empty()

        # Keep reading until the process terminates
        while process.poll() is None:
            stdout_line = process.stdout.readline()
            stderr_line = process.stderr.readline()
            if stdout_line:
                stdout_line_strip = stdout_line.strip()
                if stdout_line_strip:
                    logging.debug(f"stdout: {stdout_line_strip}") # Log debug level
                    full_stdout.append(stdout_line_strip)
                    stdout_area.text("\n".join(full_stdout[-20:]))
            if stderr_line:
                stderr_line_strip = stderr_line.strip()
                if stderr_line_strip:
                    logging.warning(f"stderr: {stderr_line_strip}") # Log warnings
                    full_stderr.append(stderr_line_strip)
                    stderr_area.text("\n".join(full_stderr))
            time.sleep(0.1) # Slightly longer sleep

        # Capture any remaining output after process finishes
        remaining_stdout, remaining_stderr = process.communicate()
        if remaining_stdout:
             for line in remaining_stdout.splitlines():
                 if line.strip(): full_stdout.append(line.strip())
        if remaining_stderr:
             for line in remaining_stderr.splitlines():
                 if line.strip(): full_stderr.append(line.strip())

        # Update final display
        stdout_area.text("\n".join(full_stdout))
        stderr_area.text("\n".join(full_stderr))

        returncode = process.returncode
        logging.info(f"magic-pdf process finished with return code {returncode}")

        if returncode != 0:
            st.error(f"magic-pdf command failed with return code {returncode}")
            error_message = "\n".join(full_stderr) if full_stderr else f"Exited with code {returncode}"
            logging.error(f"magic-pdf failed. Stderr: {error_message}")
            return None, error_message

        logging.info(f"Verifying results directory: {results_dir}")
        if not results_dir.is_dir():
             logging.error(f"Expected results directory NOT FOUND: {results_dir}")
             st.error(f"Expected results directory NOT FOUND: {results_dir}")
             try:
                 contents = [str(p) for p in job_output_base.iterdir()]
                 logging.info(f"Contents of base output dir {job_output_base}: {contents}")
                 st.text_area("Contents of Base Output Dir", "\n".join(contents), height=100)
             except Exception as list_e:
                 logging.error(f"Could not list contents of {job_output_base}: {list_e}")
                 st.error(f"Could not list contents of {job_output_base}: {list_e}")
             return None, f"Results directory missing after processing: {results_dir}"

        logging.info(f"Results directory verified: {results_dir}")
        st.success(f"Results directory verified.")
        return results_dir, None # Success

    except FileNotFoundError:
        err_msg = "Error: 'magic-pdf' command not found. Is it installed and in PATH? Did the runtime restart?"
        logging.error(err_msg, exc_info=True)
        st.error(err_msg)
        return None, "magic-pdf command not found"
    except Exception as e:
        err_msg = f"An exception occurred while running magic-pdf: {e}"
        logging.error(err_msg, exc_info=True)
        st.error(err_msg)
        return None, str(e)


def get_download_link(file_path: Path, link_text: str, filename: str):
    """Generates a download link for a file."""
    if not file_path.is_file():
        logging.warning(f"Download link requested for non-existent file: {file_path}")
        return f"<i>{filename} not found</i>"
    try:
        with open(file_path, "rb") as f:
            bytes_data = f.read()
        b64 = base64.b64encode(bytes_data).decode()
        return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">{link_text}</a>'
    except Exception as e:
        logging.error(f"Error creating download link for {file_path}: {e}")
        return f"<i>Error creating link for {filename}: {e}</i>"

def create_word_doc(md_content: str, output_docx_path: Path):
    """Creates a basic Word document from Markdown text content."""
    try:
        logging.info(f"Creating Word doc at: {output_docx_path}")
        document = Document()
        for paragraph_text in md_content.split('\n'):
            if paragraph_text.strip():
                 document.add_paragraph(paragraph_text)
        document.save(output_docx_path)
        logging.info(f"Word doc saved successfully.")
        return True
    except Exception as e:
        logging.error(f"Error creating Word document {output_docx_path}: {e}", exc_info=True)
        st.error(f"Error creating Word document: {e}")
        return False


# --- Streamlit App UI ---

st.title("📄 MinerU / Magic-PDF Processor")
st.markdown("Upload a PDF file, select the language, and process it using `magic-pdf`.")

# --- Session State Initialization ---
if 'job_id' not in st.session_state: st.session_state.job_id = None
if 'file_hash' not in st.session_state: st.session_state.file_hash = None
if 'original_stem' not in st.session_state: st.session_state.original_stem = None
if 'input_pdf_path' not in st.session_state: st.session_state.input_pdf_path = None
if 'results_dir' not in st.session_state: st.session_state.results_dir = None
if 'processing_done' not in st.session_state: st.session_state.processing_done = False
if 'error_message' not in st.session_state: st.session_state.error_message = None

# --- UI Elements ---
uploaded_file = st.file_uploader("1. Choose a PDF file", type="pdf", key="pdf_uploader")
selected_language_name = st.selectbox(
    "2. Select Language (or Auto Detect)",
    options=list(LANGUAGES.keys()),
    index=0, # Default to "Auto Detect"
    key="lang_select"
)
lang_code = LANGUAGES[selected_language_name]

process_button = st.button("✨ Process PDF", key="process_button", disabled=(uploaded_file is None), type="primary")

# --- Processing Logic ---
if process_button and uploaded_file is not None:
    # Reset state for a new run
    st.session_state.job_id = str(uuid.uuid4())
    st.session_state.results_dir = None
    st.session_state.processing_done = False
    st.session_state.error_message = None
    st.session_state.file_hash = None
    st.session_state.original_stem = None
    st.session_state.input_pdf_path = None
    logging.info(f"--- Starting New Job: {st.session_state.job_id} ---")

    # Save file, get hash and original stem
    saved_path, file_hash, original_stem = save_uploaded_file(uploaded_file)

    if saved_path and file_hash and original_stem:
        st.session_state.input_pdf_path = str(saved_path)
        st.session_state.file_hash = file_hash
        st.session_state.original_stem = original_stem
        logging.info(f"Job {st.session_state.job_id}: File Hash={file_hash}, Original Stem={original_stem}")

        job_output_base, status_file_path = get_job_dirs(file_hash)
        logging.info(f"Job {st.session_state.job_id}: Output Base={job_output_base}, Status File={status_file_path}")

        status_data = read_status(status_file_path)

        # --- Resume Logic ---
        # Check status AND if the expected results dir actually exists
        if status_data.get("status") == "completed" and status_data.get("results_dir"):
            potential_results_dir = Path(status_data["results_dir"])
            logging.info(f"Job {st.session_state.job_id}: Found previous 'completed' status. Checking results dir: {potential_results_dir}")
            if potential_results_dir.is_dir():
                st.success("Found completed results from a previous run for this file.")
                logging.info(f"Job {st.session_state.job_id}: Resuming with existing results.")
                st.session_state.results_dir = potential_results_dir
                st.session_state.processing_done = True
                st.session_state.error_message = None
                st.rerun() # Rerun to display results
            else:
                 st.warning(f"Previous status was 'completed', but results dir '{potential_results_dir}' not found. Reprocessing.")
                 logging.warning(f"Job {st.session_state.job_id}: Results dir missing despite 'completed' status. Resetting.")
                 status_data = {} # Reset status to force reprocessing
        else:
            # --- Start New Processing ---
            if status_data:
                st.info(f"Previous status was '{status_data.get('status', 'unknown')}'. Starting new processing job...")
                logging.info(f"Job {st.session_state.job_id}: Previous status was '{status_data.get('status', 'unknown')}' or incomplete. Starting fresh.")
            else:
                 st.info("No previous status found. Starting new processing job...")
                 logging.info(f"Job {st.session_state.job_id}: No status file found. Starting fresh.")

            # Write initial "started" status
            status_data = {
                "job_id": st.session_state.job_id,
                "file_hash": file_hash,
                "original_filename": uploaded_file.name,
                "original_stem": original_stem,
                "language": lang_code if lang_code else "auto",
                "status": "started",
                "results_dir": None,
                "error": None
            }
            write_status(status_file_path, status_data)

            with st.spinner(f"Processing PDF with magic-pdf (lang={lang_code or 'auto'})... This can take several minutes."):
                results_dir_path, error_msg = run_magic_pdf(saved_path, original_stem, job_output_base, lang_code)

            if results_dir_path and not error_msg:
                st.session_state.results_dir = results_dir_path
                st.session_state.processing_done = True
                st.session_state.error_message = None
                status_data["status"] = "completed"
                status_data["results_dir"] = str(results_dir_path)
                write_status(status_file_path, status_data)
                logging.info(f"Job {st.session_state.job_id}: Processing completed successfully.")
                st.rerun() # Rerun to display results
            else:
                st.session_state.error_message = error_msg or "Unknown processing error."
                status_data["status"] = "error"
                status_data["error"] = st.session_state.error_message
                write_status(status_file_path, status_data)
                logging.error(f"Job {st.session_state.job_id}: Processing failed: {st.session_state.error_message}")
                st.rerun() # Rerun to display error
    else:
         # Failed to save upload
         st.error("Failed to save the uploaded file. Please check logs or try again.")
         logging.error("Failed to save uploaded file.")
         # Reset relevant state if save fails
         st.session_state.input_pdf_path = None
         st.session_state.file_hash = None
         st.session_state.original_stem = None


# --- Display Results or Errors ---
if st.session_state.processing_done and st.session_state.results_dir:
    st.markdown("---")
    st.header("📊 Processing Results")
    results_dir = Path(st.session_state.results_dir)
    output_file_stem = st.session_state.original_stem or "output"

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Extracted Markdown")
        md_file_path = results_dir / "result.md"
        if md_file_path.is_file():
            try:
                md_content = md_file_path.read_text(encoding='utf-8')
                st.markdown(md_content, unsafe_allow_html=True)

                temp_docx_path = results_dir / f"{output_file_stem}_output.docx"
                docx_link = ""
                if create_word_doc(md_content, temp_docx_path):
                     docx_link = get_download_link(temp_docx_path, "Download Word (.docx)", f"{output_file_stem}.docx")

                md_link = get_download_link(md_file_path, "Download Markdown (.md)", f"{output_file_stem}.md")

                st.markdown("---")
                st.markdown(f"{md_link}  |  {docx_link}", unsafe_allow_html=True)

            except Exception as e:
                st.error(f"Error reading or displaying Markdown file: {e}")
                logging.error(f"Error reading/displaying results {md_file_path}: {e}", exc_info=True)
                st.code(f"Attempted to read: {md_file_path}")
        else:
            st.warning(f"Could not find result.md in {results_dir}")
            logging.warning(f"result.md not found in {results_dir}")

    with col2:
        st.subheader("🖼️ Page Analysis (Images)")
        try:
            image_files = sorted(list(results_dir.glob(f"{output_file_stem}*.png")))
            if not image_files:
                image_files = sorted(list(results_dir.glob("*.png"))) # Fallback

            if image_files:
                st.info(f"Found {len(image_files)} preview image(s).")
                tab_titles = [f"Image {i+1} ({img_path.name})" for i, img_path in enumerate(image_files)]
                if len(tab_titles) > 0:
                     image_tabs = st.tabs(tab_titles)
                     for i, img_path in enumerate(image_files):
                         with image_tabs[i]:
                            try:
                                st.image(str(img_path), use_column_width=True)
                            except Exception as e:
                                st.error(f"Could not display image {img_path.name}: {e}")
                                logging.error(f"Error displaying image {img_path}: {e}")
                else:
                     st.info("No images found to display in tabs.")
            else:
                st.info("No preview images found in the output directory.")
        except Exception as e:
             st.error(f"Error searching for images in {results_dir}: {e}")
             logging.error(f"Error finding images in {results_dir}: {e}")


    st.markdown(f"--- \n*Full output files are located in:* `{results_dir}`")
    try:
        contents = "\n".join([f.name for f in results_dir.iterdir()])
        st.expander("Full File Listing").text(contents)
    except Exception as e:
        logging.warning(f"Could not list files in {results_dir}: {e}")


elif st.session_state.error_message:
    st.markdown("---")
    st.error(f"Processing failed:")
    # Use markdown with code block for better formatting of multiline errors
    st.markdown(f"```\n{st.session_state.error_message}\n```")
    st.info("Check the stderr output expander during the run or the Colab output/logs for more details.")

# --- Footer/Info ---
st.sidebar.markdown("---")
st.sidebar.info(f"Using {'Google Drive' if DRIVE_AVAILABLE else 'Temporary Colab Storage'}")
st.sidebar.info(f"Base Path: {APP_BASE_DIR}")