"""Enkle offentlige funktioner til Microsoft Forms-svar."""

from playwright.async_api import Page

from q_ms_forms.api_client import MicrosoftForms
from q_ms_forms.models import DeleteResponsesResult


async def get_response_count(
    page: Page,
    form_id: str,
    question_id: str,
    *,
    timeout_ms: int = 30_000,
) -> int:
    """Returnerer antal svar i den angivne Microsoft Forms-formular."""
    client = MicrosoftForms(
        page=page,
        form_id=form_id,
        question_id=question_id,
        timeout_ms=timeout_ms,
    )

    return await client.get_response_count()


async def delete_current_response(
    page: Page,
    form_id: str,
    question_id: str,
    *,
    timeout_ms: int = 30_000,
) -> None:
    """Sletter permanent det svar, der aktuelt vises."""
    client = MicrosoftForms(
        page=page,
        form_id=form_id,
        question_id=question_id,
        timeout_ms=timeout_ms,
    )

    await client.delete_current_response()


async def delete_all_responses(
    page: Page,
    form_id: str,
    question_id: str,
    *,
    confirm: bool = False,
    timeout_ms: int = 30_000,
    max_deletions: int = 100,
) -> DeleteResponsesResult:
    """Sletter permanent alle svar og returnerer resultatet."""
    client = MicrosoftForms(
        page=page,
        form_id=form_id,
        question_id=question_id,
        timeout_ms=timeout_ms,
        max_deletions=max_deletions,
    )

    return await client.delete_all_responses(
        confirm=confirm,
    )


__all__ = [
    "delete_all_responses",
    "delete_current_response",
    "get_response_count",
]