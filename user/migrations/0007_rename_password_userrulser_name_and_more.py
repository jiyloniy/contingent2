# Generated by Django 5.0 on 2024-02-21 17:09

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0006_remove_organization_user_remove_userrulser_user_and_more'),
    ]

    operations = [
        migrations.RenameField(
            model_name='userrulser',
            old_name='password',
            new_name='name',
        ),
        migrations.RemoveField(
            model_name='userrulser',
            name='username',
        ),
        migrations.AddField(
            model_name='usermodel',
            name='password',
            field=models.CharField(default=1, max_length=100),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='usermodel',
            name='username',
            field=models.CharField(default=1, max_length=100),
            preserve_default=False,
        ),
        migrations.AlterField(
            model_name='usermodel',
            name='rule',
            field=models.ManyToManyField(related_name='usermodelrule', to='user.userrulser'),
        ),
    ]
