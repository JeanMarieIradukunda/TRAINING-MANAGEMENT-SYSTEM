from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0003_lessonplan'),
    ]

    operations = [
        migrations.AddField(
            model_name='module',
            name='num_terms',
            field=models.PositiveSmallIntegerField(
                default=1,
                help_text="How many terms this module's Scheme of Work is split across.",
            ),
        ),
        migrations.AddField(
            model_name='module',
            name='term_weeks',
            field=models.CharField(
                blank=True,
                max_length=100,
                help_text=(
                    "Comma-separated number of weeks for each term, in order, e.g. "
                    "'12,12,10' for a 3-term module where the last term is shorter. "
                    "Leave blank to split evenly (12 weeks per term by default)."
                ),
            ),
        ),
    ]