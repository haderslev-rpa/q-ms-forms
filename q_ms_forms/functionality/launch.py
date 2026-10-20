"""Navigation til Microsoft Forms med bagudkompatible URL-parametre."""

from __future__ import annotations

import logging
from urllib.parse import (
    parse_qsl,
    urlencode,
    urlsplit,
    urlunsplit,
)

from playwright.async_api import Page
from playwright.async_api import (
    TimeoutError as PlaywrightTimeoutError,
)

from q_haderslev_vbo.playwright.microsoft_login import (
    is_microsoft_login_url,
    login_via_microsoft,
)
from q_ms_forms.selectors import (
    FLERE_INDSTILLINGER_ARIA_LABEL,
    SVARANTAL_SELECTOR,
    TOM_FORMULAR_TEKST,
)


logger = logging.getLogger(__name__)


FORMS_BASE_URL = (
    "https://forms.cloud.microsoft/"
    "Pages/DesignPageV2.aspx"
)

NAVIGATION_TIMEOUT_MS = 120_000
ELEMENT_TIMEOUT_MS = 60_000


# --------------------------------------------------
# URL-BYGNING
# --------------------------------------------------

def build_forms_overview_url(
    form_id: str | None = None,
    *,
    forms_url: str | None = None,
) -> str:
    """Returnerer URL til svaroversigten.

    forms_url bruges uændret, når den er angivet.

    Ellers bygges den tidligere URL ud fra form_id.
    Dermed virker både den eksisterende api_client.py
    og ny kode, der arbejder med en fuld URL.
    """
    if forms_url is not None:
        return _validate_forms_url(
            forms_url
        )

    normalized_form_id = _required_text(
        form_id,
        "form_id",
    )

    query = urlencode(
        {
            "prevorigin": "shell",
            "origin": "NeoPortalPage",
            "subpage": "design",
            "id": normalized_form_id,
            "analysis": "true",
            "tab": "0",
        }
    )

    return f"{FORMS_BASE_URL}?{query}"


def build_forms_response_url(
    form_id: str | None = None,
    question_id: str | None = None,
    response_index: int = 0,
    *,
    forms_url: str | None = None,
) -> str:
    """Returnerer URL til et konkret nulbaseret svarindeks."""
    normalized_question_id = _required_text(
        question_id,
        "question_id",
    )

    if (
        isinstance(response_index, bool)
        or not isinstance(response_index, int)
        or response_index < 0
    ):
        raise ValueError(
            "response_index skal være "
            "et heltal på mindst 0."
        )

    overview_url = build_forms_overview_url(
        form_id=form_id,
        forms_url=forms_url,
    )

    parts = urlsplit(
        overview_url
    )

    query = dict(
        parse_qsl(
            parts.query,
            keep_blank_values=True,
        )
    )

    query.update(
        {
            "analysis": "true",
            "topview": "SurveyResults",
            "qid": normalized_question_id,
            "ridx": str(response_index),
        }
    )

    return urlunsplit(
        (
            parts.scheme,
            parts.netloc,
            parts.path,
            urlencode(query),
            parts.fragment,
        )
    )


# --------------------------------------------------
# PUBLIC LAUNCH-FUNKTIONER
# --------------------------------------------------

async def launch_ms_forms(
    page: Page,
    form_id: str | None = None,
    *,
    forms_url: str | None = None,
    credential_name: str | None = None,
    email: str | None = None,
    password: str | None = None,
) -> Page:
    """Åbner svaroversigten.

    Der kan angives enten:

    - form_id
    - forms_url

    Hvis Microsoft viser login, skal der samtidig være angivet:

    - credential_name

    eller:

    - email
    - password

    Output:
        Den eksisterende Playwright Page.
    """
    target_url = build_forms_overview_url(
        form_id=form_id,
        forms_url=forms_url,
    )

    await _goto(
        page=page,
        url=target_url,
    )

    await _login_if_needed(
        page=page,
        target_url=target_url,
        credential_name=credential_name,
        email=email,
        password=password,
    )

    await _wait_for_forms_page(
        page=page,
    )

    logger.info(
        "Microsoft Forms-svaroversigten "
        "er klar. URL: %s",
        page.url,
    )

    return page


async def launch_ms_forms_response(
    page: Page,
    form_id: str | None = None,
    question_id: str | None = None,
    response_index: int = 0,
    *,
    forms_url: str | None = None,
    credential_name: str | None = None,
    email: str | None = None,
    password: str | None = None,
) -> Page:
    """Åbner et konkret Microsoft Forms-svar.

    Der kan angives enten:

    - form_id
    - forms_url

    Output:
        Den eksisterende Playwright Page.
    """
    target_url = build_forms_response_url(
        form_id=form_id,
        question_id=question_id,
        response_index=response_index,
        forms_url=forms_url,
    )

    await _goto(
        page=page,
        url=target_url,
    )

    await _login_if_needed(
        page=page,
        target_url=target_url,
        credential_name=credential_name,
        email=email,
        password=password,
    )

    await _wait_for_forms_page(
        page=page,
    )

    logger.info(
        "Microsoft Forms-svaret "
        "er klar. URL: %s",
        page.url,
    )

    return page


# --------------------------------------------------
# LOGIN
# --------------------------------------------------

async def _login_if_needed(
    page: Page,
    target_url: str,
    credential_name: str | None,
    email: str | None,
    password: str | None,
) -> None:
    """Logger ind, hvis Microsoft viser login.

    Efter login åbnes den oprindelige mål-URL igen.
    """
    if not is_microsoft_login_url(
        page
    ):
        return

    has_direct_login = bool(
        email
        and password
    )

    has_credential = bool(
        isinstance(
            credential_name,
            str,
        )
        and credential_name.strip()
    )

    if (
        not has_direct_login
        and not has_credential
    ):
        raise RuntimeError(
            "Microsoft-login er påkrævet, "
            "men launchfunktionen modtog "
            "hverken credential_name "
            "eller email/password."
        )

    await login_via_microsoft(
        page=page,
        credential_name=credential_name,
        email=email,
        password=password,
        action_timeout_ms=ELEMENT_TIMEOUT_MS,
        login_timeout_ms=NAVIGATION_TIMEOUT_MS,
    )

    await _goto(
        page=page,
        url=target_url,
    )

    if is_microsoft_login_url(
        page
    ):
        raise RuntimeError(
            "Browseren står stadig på "
            "Microsoft-login-siden efter login. "
            f"Aktuel URL: {page.url}"
        )


# --------------------------------------------------
# NAVIGATION OG VENTETID
# --------------------------------------------------

async def _goto(
    page: Page,
    url: str,
) -> None:
    """Navigerer til URL'en og venter på DOMContentLoaded."""
    await page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=NAVIGATION_TIMEOUT_MS,
    )


async def _wait_for_forms_page(
    page: Page,
) -> None:
    """Venter på at Microsoft Forms er klar.

    Funktionen accepterer én af følgende tilstande:

    - element med svarantal
    - svarmenuen
    - Forms' tombesked
    """
    try:
        await page.wait_for_function(
            """
            ([countSelector, emptyText, optionsLabel]) => {
                const bodyText =
                    document.body?.innerText || "";

                if (bodyText.includes(emptyText)) {
                    return true;
                }

                if (
                    document.querySelector(
                        countSelector
                    )
                ) {
                    return true;
                }

                return Array.from(
                    document.querySelectorAll(
                        "button"
                    )
                ).some((button) => {
                    return (
                        button.getAttribute(
                            "aria-label"
                        ) === optionsLabel
                        && button.offsetParent !== null
                    );
                });
            }
            """,
            arg=[
                SVARANTAL_SELECTOR,
                TOM_FORMULAR_TEKST,
                FLERE_INDSTILLINGER_ARIA_LABEL,
            ],
            timeout=ELEMENT_TIMEOUT_MS,
        )

    except PlaywrightTimeoutError as error:
        raise RuntimeError(
            "Microsoft Forms blev åbnet, "
            "men siden blev ikke klar. "
            f"Aktuel URL: {page.url}"
        ) from error


# --------------------------------------------------
# VALIDERING
# --------------------------------------------------

def _validate_forms_url(
    value: str,
) -> str:
    """Validerer en fuld Microsoft Forms-URL."""
    normalized = _required_text(
        value,
        "forms_url",
    )

    parts = urlsplit(
        normalized
    )

    if (
        parts.scheme != "https"
        or parts.netloc.casefold()
        != "forms.cloud.microsoft"
    ):
        raise ValueError(
            "forms_url skal være en HTTPS-URL "
            "på forms.cloud.microsoft."
        )

    return normalized


def _required_text(
    value: str | None,
    field_name: str,
) -> str:
    """Returnerer valideret, ikke-tom tekst."""
    if (
        not isinstance(value, str)
        or not value.strip()
    ):
        raise ValueError(
            f"{field_name} skal være "
            "en ikke-tom tekst."
        )

    return value.strip()


__all__ = [
    "build_forms_overview_url",
    "build_forms_response_url",
    "launch_ms_forms",
    "launch_ms_forms_response",
]