# Generated by Django 5.0 on 2024-03-04 20:18

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0003_alter_guruh_bosqich'),
    ]

    operations = [
        migrations.AddField(
            model_name='organization',
            name='full_name',
            field=models.CharField(blank=True, default='Sining universitetingiz nomi', max_length=100, null=True),
        ),
    ]
