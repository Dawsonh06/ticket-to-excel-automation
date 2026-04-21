"""
function_app.py
===============
Azure Functions entry point for the ticket processing automation.
Runs on a timer trigger and delegates to ticket_processor.main().
"""

import logging

import azure.functions as func

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)


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
    from ticket_processor import main
    main()
