from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0005_alter_logo_image_textfield'),
    ]

    operations = [
        migrations.AddField(
            model_name='module',
            name='is_active',
            field=models.BooleanField(
                default=True,
                help_text="Unlocked (usable) modules appear in the public generators. "
                          "Lock a module to block it from being used there.",
            ),
        ),
        migrations.AlterModelOptions(
            name='module',
            options={
                'ordering': ['mod_code'],
                'permissions': [('toggle_module_status', 'Can lock/unlock module usage')],
            },
        ),
    ]
