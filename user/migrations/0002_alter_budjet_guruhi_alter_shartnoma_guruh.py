# Generated by Django 5.0 on 2024-03-03 17:06

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0001_initial'),
    ]

    operations = [
        migrations.AlterField(
            model_name='budjet',
            name='guruhi',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='guruhbudjet', to='user.guruh'),
        ),
        migrations.AlterField(
            model_name='shartnoma',
            name='guruh',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='guruhshartnoma', to='user.guruh'),
        ),
    ]
