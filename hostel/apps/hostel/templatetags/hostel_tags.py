from django import template

register = template.Library()


@register.filter
def get_item(dictionary, key):
    """{{ dict|get_item:key }} — safe dict lookup in templates."""
    if isinstance(dictionary, dict):
        return dictionary.get(key)
    return None


@register.filter
def split_items(value):
    """Split a menu string (comma or newline separated) into a clean list."""
    if not value:
        return []
    items = []
    for part in value.replace('\n', ',').split(','):
        s = part.strip()
        if s:
            items.append(s)
    return items
