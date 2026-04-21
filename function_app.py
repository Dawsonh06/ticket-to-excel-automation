"""
function_app.py
===============
Azure Functions entry point for the ticket processing automation.
Runs on a timer trigger and delegates to ticket_processor.main().
"""

import logging
import subprocess
import sys

import azure.functions as func

app = func.FunctionApp()


def ensure_dependencies():
    """Install system packages required by pytesseract and zxingcpp if missing."""
    try:
        subprocess.run(
            ["tesseract", "--version"],
            capture_output=True,
            check=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        logging.warning(
            "tesseract not found — attempting apt-get install. "
            "This should not happen in a correctly provisioned environment."
        )
        subprocess.run(
            ["apt-get", "install", "-y", "tesseract-ocr", "libzbar0"],
            capture_output=True,
        )


# Timer schedule: every 5 minutes.  Adjust the cron expression in host.json or
# here if a different interval is needed.
@app.timer_trigger(
    schedule="0 */5 * * * *",
    arg_name="timer",
    run_on_startup=False,
    use_monitor=False,
)
def ticket_processor_timer(timer: func.TimerRequest) -> None:
    """Azure Functions timer trigger — runs ticket_processor.main()."""
    ensure_dependencies()

    import ticket_processor
    ticket_processor.main()
