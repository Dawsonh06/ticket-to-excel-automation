"""
function_app.py
===============
Azure Functions entry point for the ticket processing automation.
Runs on a timer trigger and delegates to ticket_processor.main().
"""

import logging
import subprocess

import azure.functions as func

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


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


@app.timer_trigger(
    schedule="0 0 0 * * *",
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=False,
)
def ticket_processor_timer(myTimer: func.TimerRequest) -> None:
    if myTimer.past_due:
        logging.info("Timer is past due")
    logging.info("Ticket processor timer triggered")
    ensure_dependencies()
    from ticket_processor import main
    main()
