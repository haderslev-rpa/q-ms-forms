"""Resultatmodeller til q-ms-forms."""

from dataclasses import asdict, dataclass


@dataclass(frozen=True, slots=True)
class DeleteResponsesResult:
    """
    Resultat fra permanent sletning af Microsoft Forms-svar.

    Felter:
        count_before:
            Antal svar før sletningen.

        count_deleted:
            Antal svar, som blev slettet.

        count_after:
            Antal svar efter sletningen.
    """

    count_before: int
    count_deleted: int
    count_after: int

    @property
    def all_deleted(self) -> bool:
        """
        Returnerer True, når alle registrerede svar blev slettet.
        """
        return (
            self.count_after == 0
            and self.count_deleted == self.count_before
        )

    def to_dict(self) -> dict:
        """
        Returnerer resultatet som en almindelig dictionary.
        """
        result = asdict(self)
        result["all_deleted"] = self.all_deleted
        return result