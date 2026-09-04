from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('core', '0006_module_is_active_and_toggle_permission'),
    ]

    operations = [
        migrations.AddField(
            model_name='trainer',
            name='user',
            field=models.OneToOneField(
                blank=True,
                help_text="The login account this trainer uses to sign in. "
                          "Leave blank if this trainer doesn't need portal access.",
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name='trainer_profile',
                to=settings.AUTH_USER_MODEL,
            ),
        ),
    ]
