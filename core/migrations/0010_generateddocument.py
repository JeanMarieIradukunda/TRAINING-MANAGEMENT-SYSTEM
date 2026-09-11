from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('core', '0009_traineraccess'),
    ]

    operations = [
        migrations.CreateModel(
            name='GeneratedDocument',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('doc_type', models.CharField(choices=[
                    ('scheme_of_work', 'Scheme of Work'),
                    ('lesson_plan', 'Lesson Plan'),
                    ('assessment_plan', 'Assessment Plan'),
                ], max_length=20)),
                ('title', models.CharField(blank=True, max_length=255)),
                ('meta_snapshot', models.JSONField(blank=True, default=list)),
                ('filename', models.CharField(max_length=255)),
                ('content_type', models.CharField(max_length=150)),
                ('file_data', models.BinaryField()),
                ('file_size', models.PositiveIntegerField(default=0)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('generated_by', models.ForeignKey(
                    blank=True, null=True,
                    on_delete=django.db.models.deletion.SET_NULL,
                    related_name='generated_documents',
                    to=settings.AUTH_USER_MODEL,
                )),
                ('trainer', models.ForeignKey(
                    blank=True, null=True,
                    on_delete=django.db.models.deletion.SET_NULL,
                    related_name='generated_documents',
                    to='core.trainer',
                )),
            ],
            options={
                'verbose_name': 'Generated document',
                'ordering': ['-created_at'],
            },
        ),
    ]
