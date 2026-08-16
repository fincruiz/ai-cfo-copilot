from __future__ import annotations

FINANCE_WRITE_ROLES = frozenset({
    "owner",
    "admin",
    "cfo",
    "finance_manager",
    "accountant",
})

COMPANY_ADMIN_ROLES = frozenset({"owner", "admin"})


def role_value(role: object) -> str:
    value = getattr(role, "value", role)
    return str(value)


def can_finance_write(role: object) -> bool:
    return role_value(role) in FINANCE_WRITE_ROLES


def can_company_admin(role: object) -> bool:
    return role_value(role) in COMPANY_ADMIN_ROLES
