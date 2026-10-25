"""Domain errors for docx-cli."""


class DocxCreatorError(Exception):
    """Base exception."""


class SpecError(DocxCreatorError):
    """Input spec violates the machine contract."""


class BuildError(DocxCreatorError):
    """Generation failed mid-way."""


class ValidationError(DocxCreatorError):
    """Generated document does not conform to the profile/contract."""


class ProfileError(DocxCreatorError):
    """Requested profile is unknown or misconfigured."""
