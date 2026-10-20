"""Valideringsfunktioner til q-ms-forms."""


def require_text(value: str, field_name: str) -> str:
    """Returnerer en valideret, ikke-tom tekstværdi."""
    if not isinstance(value, str):
        raise TypeError(f"{field_name} skal være tekst.")

    cleaned_value = value.strip()
    if not cleaned_value:
        raise ValueError(f"{field_name} må ikke være tom.")

    return cleaned_value


def require_positive_int(value: int, field_name: str) -> int:
    """Returnerer et valideret positivt heltal."""
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{field_name} skal være et helt tal.")

    if value <= 0:
        raise ValueError(f"{field_name} skal være større end 0.")

    return value


def require_non_negative_int(value: int, field_name: str) -> int:
    """Returnerer et valideret heltal på mindst 0."""
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{field_name} skal være et helt tal.")

    if value < 0:
        raise ValueError(f"{field_name} skal være mindst 0.")

    return value
