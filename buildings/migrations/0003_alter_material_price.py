# Generated by Django 5.1.7 on 2025-04-26 19:29

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('buildings', '0002_alter_worker_salary'),
    ]

    operations = [
        migrations.AlterField(
            model_name='material',
            name='price',
            field=models.FloatField(blank=True, null=True),
        ),
    ]
