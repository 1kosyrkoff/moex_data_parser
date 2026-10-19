FROM python:3.13-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /app
COPY requirements-service.lock ./
RUN pip install --no-cache-dir -r requirements-service.lock \
    && useradd --uid 10001 --create-home bond
# Deliberate allowlist: notebooks and unrelated report scripts are not imported.
COPY moex_bonds.py moex_cls.py ./
COPY bond_service ./bond_service
ARG SOURCE_REVISION=unknown
ENV SOURCE_REVISION=$SOURCE_REVISION
USER bond
CMD ["python", "-m", "bond_service.scheduler"]
