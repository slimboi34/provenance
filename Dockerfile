FROM python:3.12-slim
WORKDIR /app
COPY pyproject.toml README.md ./
COPY provenance ./provenance
RUN pip install --no-cache-dir ".[mcp]" formulas openpyxl && useradd --create-home app
USER app
ENV PYTHONUNBUFFERED=1
EXPOSE 8000
CMD ["provenance-mcp", "--http"]
