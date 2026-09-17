from django import template

from core.models import user_is_dos

register = template.Library()


@register.filter(name='is_dos')
def is_dos(user):
    """Template usage: {% if request.user|is_dos %} ... {% endif %}"""
    return user_is_dos(user)


@register.filter(name='display_name')
def display_name(user):
    """
    The person's real name for header/greeting display, instead of their
    login username. Tries, in order:

      1. Trainer.full_name (fname + lname) for a Trainer-linked login -
         always filled in, since both fields are required when a Trainer
         record is created.
      2. auth.User.get_full_name() (first_name + last_name) for an Admin
         or Dean of Studies login, when an admin has filled those in from
         the Django admin "Users" screen.
      3. The username, title-cased, if neither of the above is set - so
         the header never shows a blank name.

    Safe to call with an AnonymousUser (returns '').
    """
    if not getattr(user, 'is_authenticated', False):
        return ''

    trainer = getattr(user, 'trainer_profile', None)
    if trainer is not None:
        name = (trainer.full_name or '').strip()
        if name:
            return name

    full_name = (user.get_full_name() or '').strip()
    if full_name:
        return full_name

    return user.get_username().title()