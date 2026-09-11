from django import template

from core.models import user_is_dos

register = template.Library()


@register.filter(name='is_dos')
def is_dos(user):
    """Template usage: {% if request.user|is_dos %} ... {% endif %}"""
    return user_is_dos(user)
