"""
Custom authentication backend for Trainer logins.

Why this exists
----------------
Django's login form (AuthenticationForm) always calls
django.contrib.auth.authenticate(), which — with no custom backend
configured — only ever checks the built-in auth_user table via
ModelBackend. That's correct for the admin/superuser login, but it meant
Trainer.username / Trainer.password_hash (on core_trainer) were being
set by TrainerForm and then never actually read by anything: trainers
could only log in if they *also* had a separate auth_user row manually
linked through Trainer.user.

TrainerBackend below makes core_trainer the real source of truth for
trainer credentials: it checks username + password against
Trainer.username / Trainer.password_hash directly. Django's session
framework still needs a concrete auth.User instance to attach the
session to, so on success this backend returns the trainer's linked
`Trainer.user` (a lightweight, permission-less shadow account that
TrainerForm.save() creates/keeps in sync automatically — admins never
manage it by hand).

Registered in settings.AUTHENTICATION_BACKENDS *after* ModelBackend, so:
  - Admin/superuser logins keep working exactly as before, via ModelBackend
    against auth_user (tried first).
  - Only if that fails does Django try TrainerBackend against core_trainer.
"""
from django.contrib.auth.backends import ModelBackend
from django.contrib.auth.hashers import check_password

from .models import Trainer


class TrainerBackend(ModelBackend):
    def authenticate(self, request, username=None, password=None, **kwargs):
        if not username or not password:
            return None

        try:
            trainer = Trainer.objects.select_related('user').get(username=username)
        except Trainer.DoesNotExist:
            return None

        if not trainer.password_hash:
            return None

        if not check_password(password, trainer.password_hash):
            return None

        user = trainer.user
        if user is None or not user.is_active:
            # No linked login account (or it's been deactivated) - nothing
            # for Django's session to attach to, so authentication fails
            # even though the trainer credentials themselves were correct.
            return None

        return user

    def get_user(self, user_id):
        # Reuse ModelBackend's implementation: once authenticate() returns
        # a real auth.User, session lookups on subsequent requests work
        # exactly like any other Django login.
        return super().get_user(user_id)
