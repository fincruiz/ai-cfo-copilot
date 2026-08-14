from collections.abc import AsyncGenerator

from sqlalchemy.ext.asyncio import (
    AsyncEngine,
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)
from sqlalchemy.pool import NullPool

from app.core.config import settings


engine: AsyncEngine | None = None
AsyncSessionLocal: async_sessionmaker[AsyncSession] | None = None

if settings.database_url:
    engine = create_async_engine(
        settings.database_url,
        echo=False,
        pool_pre_ping=True,
        connect_args={
            "statement_cache_size": 0,
            "prepared_statement_cache_size": 0,
        },
        poolclass=NullPool,
    )

    AsyncSessionLocal = async_sessionmaker(
        bind=engine,
        class_=AsyncSession,
        expire_on_commit=False,
        autoflush=False,
        autocommit=False,
    )

# Compatibility alias in case existing code uses this name.
AsyncSessionFactory = AsyncSessionLocal


async def get_db_session() -> AsyncGenerator[AsyncSession, None]:
    if AsyncSessionLocal is None:
        raise RuntimeError(
            "DATABASE_URL is not configured. Set it before using database-backed endpoints."
        )

    async with AsyncSessionLocal() as session:
        try:
            yield session
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()
