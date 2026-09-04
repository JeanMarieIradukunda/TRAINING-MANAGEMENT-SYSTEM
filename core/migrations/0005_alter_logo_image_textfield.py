from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0004_module_num_terms_module_term_weeks'),
    ]

    operations = [
        migrations.AlterField(
            model_name='logo',
            name='image',
            field=models.TextField(help_text='Base64 string or image URL'),
        ),
    ]
