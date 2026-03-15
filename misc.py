import re


def format_personnel(raw: str) -> str:
    """
    Підтримує формати:
        майор ІВАНЕНКО І.І.
        майор ІВАНЕНКО І.І
        ІВАНЕНКО І.І.
        ІВАНЕНКО І.І
    Результат: «майор ІВАНЕНКО, ІВАНЕНКО»
    """
    results = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        # Видаляємо ініціали в кінці: І.І. або І.І або І.
        cleaned = re.sub(r'\s+[А-ЯІЇЄҐA-Z]\.[А-ЯІЇЄҐA-Z]?\.?\s*$', '', line).strip()
        if cleaned:
            results.append(cleaned)
    return ', '.join(results)
