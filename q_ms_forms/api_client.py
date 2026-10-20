"""Microsoft Forms-object med synkron sletning i Forms og den delte Excel-fil."""

from __future__ import annotations

import logging
import re
from time import monotonic

from playwright.async_api import Frame, Locator, Page
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from q_ms_forms.excel_responses import ExcelResponse, FormsExcelWorkbook
from q_ms_forms.functionality.launch import (
    build_forms_overview_url,
    build_forms_response_url,
    launch_ms_forms,
    launch_ms_forms_response,
)
from q_ms_forms.models import DeleteResponsesResult
from q_ms_forms.selectors import (
    FLERE_INDSTILLINGER_ARIA_LABEL,
    SLET_BEKRAEFTELSE_SELECTOR,
    SLET_SVAR_TEKST,
    SVARANTAL_SELECTOR,
    TOM_FORMULAR_TEKST,
)
from q_ms_forms.utils import require_positive_int, require_text

logger = logging.getLogger(__name__)


class MicrosoftForms:
    """Object til læsning og sletning af Microsoft Forms-svar."""

    def __init__(
        self,
        page: Page,
        form_id: str,
        question_id: str,
        *,
        timeout_ms: int = 30_000,
        max_deletions: int = 100,
        excel_share_url: str | None = None,
        excel_sheet_name: str = "Sheet1",
        excel_id_column: str = "Id",
        excel_credential_name: str = "API_SHAREPOINT",
    ) -> None:
        if page is None or page.is_closed():
            raise ValueError("En åben Playwright Page skal angives.")
        self.page = page
        self.form_id = require_text(form_id, "form_id")
        self.question_id = require_text(question_id, "question_id")
        self.timeout_ms = require_positive_int(timeout_ms, "timeout_ms")
        self.max_deletions = require_positive_int(max_deletions, "max_deletions")
        self.excel_workbook = (
            FormsExcelWorkbook(
                share_url=excel_share_url,
                sheet_name=excel_sheet_name,
                id_column=excel_id_column,
                credential_name=excel_credential_name,
            )
            if excel_share_url
            else None
        )

    @property
    def overview_url(self) -> str:
        return build_forms_overview_url(form_id=self.form_id)

    def response_url(self, response_index: int = 0) -> str:
        self._validate_response_index(response_index)
        return build_forms_response_url(
            form_id=self.form_id,
            question_id=self.question_id,
            response_index=response_index,
        )

    async def open_overview(self) -> None:
        self._ensure_page_open()
        await launch_ms_forms(page=self.page, form_id=self.form_id)

    async def open_response(self, response_index: int = 0) -> None:
        self._ensure_page_open()
        self._validate_response_index(response_index)
        await launch_ms_forms_response(
            page=self.page,
            form_id=self.form_id,
            question_id=self.question_id,
            response_index=response_index,
        )

    async def open_first_response(self) -> None:
        await self.open_response(response_index=0)

    async def get_response_count(self) -> int:
        await self.open_overview()
        groups = self.page.locator(SVARANTAL_SELECTOR)
        try:
            await groups.first.wait_for(state="visible", timeout=self.timeout_ms)
        except PlaywrightTimeoutError as error:
            if await self.has_no_responses():
                return 0
            raise RuntimeError("Kunne ikke aflæse antal Forms-svar.") from error

        for index in range(await groups.count()):
            label = await groups.nth(index).get_attribute("aria-label")
            match = re.fullmatch(
                r"\s*(\d+)\s+svar\s*",
                label or "",
                flags=re.IGNORECASE,
            )
            if match:
                return int(match.group(1))
        raise RuntimeError("Forms-elementet med antal svar blev ikke fundet.")

    async def has_no_responses(self) -> bool:
        message = self.page.get_by_text(TOM_FORMULAR_TEKST, exact=True)
        return await message.count() > 0 and await message.first.is_visible()

    async def select_response_by_id(
        self,
        response_id: int,
    ) -> None:
        """Vælger og verificerer et Forms-svar via feltet Svarperson.

        Parametre:
            response_id:
                Det stabile Id fra Excel-kolonnen Id.

        Output:
            Funktionen returnerer ikke en værdi.
            Ved succes viser Forms det samme Id som Excel.
        """
        self._ensure_page_open()
        response_id = self._validate_response_id(response_id)
        response_input = await self._find_response_id_input()

        try:
            await response_input.fill(
                str(response_id),
                timeout=self.timeout_ms,
            )
            await response_input.press(
                "Enter",
                timeout=self.timeout_ms,
            )
        except PlaywrightTimeoutError as error:
            raise RuntimeError(
                "Response-id kunne ikke indsættes i feltet Svarperson."
            ) from error

        await self.page.wait_for_timeout(750)

        # Forms kan erstatte inputelementet efter Enter.
        response_input = await self._find_response_id_input()
        actual_id = (await response_input.input_value()).strip()

        if actual_id != str(response_id):
            raise RuntimeError(
                f"Forms viser response-id {actual_id!r}, men Excel-id "
                f"{response_id!r} var forventet. Der blev ikke slettet noget."
            )

        logger.info(
            "Forms står nu på response-id %s.",
            response_id,
        )

    async def _find_response_id_input(self) -> Locator:
        """Finder et synligt Svarperson-input på tværs af frames.

        Microsoft Forms kan ændre inputfeltets role og type. Aria-label
        bruges derfor som den stabile identifikation.

        Output:
            En synlig og enabled Playwright Locator.
        """
        deadline = monotonic() + self.timeout_ms / 1000

        selectors = (
            'input[aria-label="Svarperson"]',
            '[aria-label="Svarperson"][role="spinbutton"]',
            '[aria-label="Svarperson"][role="textbox"]',
        )

        while monotonic() < deadline:
            self._ensure_page_open()

            for frame in self.page.frames:
                candidate = await self._find_response_id_input_in_frame(
                    frame=frame,
                    selectors=selectors,
                )
                if candidate is not None:
                    logger.info(
                        "Fandt feltet Svarperson i frame: %s",
                        frame.url,
                    )
                    return candidate

            await self.page.wait_for_timeout(250)

        raise RuntimeError(
            "Feltet Svarperson blev ikke fundet på Microsoft Forms-svarvisningen. "
            f"Aktuel URL: {self.page.url}. Antal frames: {len(self.page.frames)}."
        )

    async def _find_response_id_input_in_frame(
        self,
        frame: Frame,
        selectors: tuple[str, ...],
    ) -> Locator | None:
        """Finder første synlige Svarperson-input i én frame."""
        for selector in selectors:
            candidates = frame.locator(selector)

            try:
                candidate_count = await candidates.count()
            except Exception:
                continue

            for index in range(candidate_count):
                candidate = candidates.nth(index)

                try:
                    if not await candidate.is_visible():
                        continue
                    if not await candidate.is_enabled():
                        continue

                    aria_label = await candidate.get_attribute("aria-label")
                    if (
                        aria_label
                        and aria_label.strip().casefold() == "svarperson"
                    ):
                        return candidate
                except Exception:
                    continue

        return None

    async def delete_current_response(self) -> None:
        """Sletter det svar, der aktuelt vises i Forms."""
        if await self.has_no_responses():
            raise RuntimeError("Formularen har ingen svar.")

        options = await self._visible(
            self.page.get_by_role(
                "button",
                name=FLERE_INDSTILLINGER_ARIA_LABEL,
                exact=True,
            ),
            "Flere indstillinger for svar",
        )
        await options.click(timeout=self.timeout_ms)

        delete_item = await self._visible(
            self.page.get_by_text(SLET_SVAR_TEKST, exact=True),
            "Slet svar",
        )
        await delete_item.click(timeout=self.timeout_ms)

        confirm = await self._visible(
            self.page.locator(SLET_BEKRAEFTELSE_SELECTOR),
            "Bekræftelsesknappen Slet",
        )
        await confirm.click(timeout=self.timeout_ms)
        await confirm.wait_for(state="hidden", timeout=self.timeout_ms)
        await self._wait_for_response_or_empty()

    async def delete_response(self, response_index: int) -> None:
        self._validate_response_index(response_index)
        await self.open_response(response_index=response_index)
        await self.delete_current_response()

    def get_excel_responses(self) -> list[ExcelResponse]:
        """Henter og logger alle rækker fra online-Excel-filen."""
        return self._excel().get_responses()

    def get_excel_response_by_id(self, response_id: int) -> ExcelResponse | None:
        return self._excel().get_response_by_id(response_id)

    async def delete_response_and_excel_row(
        self,
        response_id: int,
    ) -> ExcelResponse:
        """Sletter Forms-svar og derefter Excel-rækken med samme Id."""
        response_id = self._validate_response_id(response_id)
        workbook = self._excel()
        excel_response = workbook.get_response_by_id(response_id)
        if excel_response is None:
            raise RuntimeError(
                f"Excel indeholder ingen række med Id {response_id}. "
                "Forms-svaret blev ikke slettet."
            )

        await self.select_response_by_id(response_id)
        await self.delete_current_response()

        try:
            return workbook.delete_response_by_id(response_id)
        except Exception as error:
            logger.critical(
                "Forms-svar %s blev slettet, men Excel-rækken blev ikke slettet.",
                response_id,
                exc_info=True,
            )
            raise RuntimeError(
                f"Forms-svar {response_id} er slettet, men Excel-oprydningen fejlede."
            ) from error

    async def delete_all_responses(self, *, confirm: bool = False) -> DeleteResponsesResult:
        """Sletter alle Forms-svar. Excel-sletning kræver id-baseret behandling."""
        if confirm is not True:
            raise ValueError("Permanent sletning kræver confirm=True.")
        count_before = await self.get_response_count()
        if count_before > self.max_deletions:
            raise RuntimeError("Antal svar overstiger max_deletions.")
        count_deleted = 0
        while count_deleted < count_before:
            await self.open_first_response()
            await self.delete_current_response()
            count_deleted += 1
        count_after = await self.get_response_count()
        return DeleteResponsesResult(
            count_before=count_before,
            count_deleted=count_deleted,
            count_after=count_after,
        )

    async def _visible(self, locator: Locator, description: str) -> Locator:
        try:
            await locator.first.wait_for(state="visible", timeout=self.timeout_ms)
            return locator.first
        except PlaywrightTimeoutError as error:
            raise RuntimeError(f"{description} blev ikke synlig.") from error

    async def _wait_for_response_or_empty(self) -> None:
        await self.page.wait_for_function(
            """
            ([emptyText, optionsLabel]) => {
                const text = document.body?.innerText || "";
                if (text.includes(emptyText)) return true;
                return Array.from(document.querySelectorAll("button")).some(
                    button => button.getAttribute("aria-label") === optionsLabel
                        && button.offsetParent !== null
                );
            }
            """,
            arg=[TOM_FORMULAR_TEKST, FLERE_INDSTILLINGER_ARIA_LABEL],
            timeout=self.timeout_ms,
        )

    def _excel(self) -> FormsExcelWorkbook:
        if self.excel_workbook is None:
            raise RuntimeError("Angiv excel_share_url ved oprettelse af MicrosoftForms.")
        return self.excel_workbook

    def _ensure_page_open(self) -> None:
        if self.page.is_closed():
            raise RuntimeError("Playwright-siden er lukket.")

    @staticmethod
    def _validate_response_index(response_index: int) -> None:
        if isinstance(response_index, bool) or not isinstance(response_index, int):
            raise TypeError("response_index skal være et heltal.")
        if response_index < 0:
            raise ValueError("response_index skal være mindst 0.")

    @staticmethod
    def _validate_response_id(response_id: int) -> int:
        if isinstance(response_id, bool) or not isinstance(response_id, int):
            raise TypeError("response_id skal være et heltal.")
        if response_id <= 0:
            raise ValueError("response_id skal være større end 0.")
        return response_id


__all__ = ["MicrosoftForms"]
