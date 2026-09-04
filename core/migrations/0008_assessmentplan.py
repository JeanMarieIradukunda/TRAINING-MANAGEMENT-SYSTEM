import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0007_trainer_user'),
    ]

    operations = [
        migrations.CreateModel(
            name='AssessmentPlan',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('module_type', models.CharField(blank=True, choices=[('core', 'Core'), ('optional', 'Optional'), ('elective', 'Elective')], max_length=20)),
                ('assessment_type', models.CharField(blank=True, choices=[('written', 'Written assessment'), ('oral', 'Oral assessment'), ('practical', 'Practical assessment'), ('assignment', 'Assignment')], max_length=20)),
                ('num_candidates', models.PositiveIntegerField(default=0)),
                ('num_invigilators', models.PositiveIntegerField(default=0)),
                ('assessment_date', models.DateField(blank=True, null=True)),
                ('resources', models.TextField(blank=True)),
                ('place', models.CharField(blank=True, max_length=150)),
                ('publication_date', models.DateField(blank=True, null=True)),
                ('observation', models.TextField(blank=True)),
                ('module', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='assessment_plans', to='core.module')),
                ('learning_outcome', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='assessment_plans', to='core.learningoutcome')),
                ('trainer', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='assessment_plans', to='core.trainer')),
            ],
            options={
                'ordering': ['-assessment_date', 'module__mod_code'],
            },
        ),
    ]
