from typing import Any


class ApplicationError(Exception):
    def __init__(
        self,
        *,
        message: str,
        error_code: str,
        status_code: int = 400,
        details: Any = None,
    ) -> None:
        self.message = message
        self.error_code = error_code
        self.status_code = status_code
        self.details = details

        super().__init__(message)


class ResourceNotFoundError(ApplicationError):
    def __init__(
        self,
        *,
        resource_name: str,
        resource_id: str | None = None,
    ) -> None:
        details = None

        if resource_id is not None:
            details = {
                "resource": resource_name,
                "resource_id": resource_id,
            }

        super().__init__(
            message=f"{resource_name} not found.",
            error_code=f"{resource_name.upper()}_NOT_FOUND",
            status_code=404,
            details=details,
        )


class ConflictError(ApplicationError):
    def __init__(
        self,
        *,
        message: str,
        error_code: str = "RESOURCE_CONFLICT",
        details: Any = None,
    ) -> None:
        super().__init__(
            message=message,
            error_code=error_code,
            status_code=409,
            details=details,
        )


class DatabaseOperationError(ApplicationError):
    def __init__(
        self,
        *,
        message: str = "A database operation failed.",
        details: Any = None,
    ) -> None:
        super().__init__(
            message=message,
            error_code="DATABASE_OPERATION_FAILED",
            status_code=500,
            details=details,
        )