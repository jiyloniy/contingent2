# Generated by Django 5.0 on 2024-03-03 14:40

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0008_remove_userrulser_organization_remove_faculty_user_and_more'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.AddField(
            model_name='yonalish',
            name='code',
            field=models.CharField(blank=True, max_length=100, null=True),
        ),
        migrations.AddField(
            model_name='yonalish',
            name='mutahasislik_2',
            field=models.BooleanField(default=False),
        ),
        migrations.CreateModel(
            name='UserRules',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('movqesi', models.CharField(max_length=100)),
                ('password', models.CharField(blank=True, max_length=100, null=True)),
                ('username', models.CharField(blank=True, max_length=100, null=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('can_add_faculty', models.BooleanField(default=False)),
                ('can_add_yonalish', models.BooleanField(default=False)),
                ('can_add_guruh', models.BooleanField(default=False)),
                ('can_update_faculty', models.BooleanField(default=False)),
                ('can_update_yonalish', models.BooleanField(default=False)),
                ('can_update_guruh', models.BooleanField(default=False)),
                ('can_delete_faculty', models.BooleanField(default=False)),
                ('can_delete_yonalish', models.BooleanField(default=False)),
                ('can_delete_guruh', models.BooleanField(default=False)),
                ('can_view_faculty', models.BooleanField(default=False)),
                ('can_view_yonalish', models.BooleanField(default=False)),
                ('can_view_guruh', models.BooleanField(default=False)),
                ('full_access', models.BooleanField(default=False)),
                ('can_delete_user', models.BooleanField(default=False)),
                ('can_add_user', models.BooleanField(default=False)),
                ('can_update_user', models.BooleanField(default=False)),
                ('can_view_user', models.BooleanField(default=False)),
                ('org', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, to='user.organization')),
                ('user', models.OneToOneField(on_delete=django.db.models.deletion.CASCADE, to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'verbose_name_plural': 'Foydalanuvchilar',
            },
        ),
    ]