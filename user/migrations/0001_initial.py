# Generated by Django 5.0 on 2024-03-03 15:50

import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name='Guruh',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('kurs', models.IntegerField(choices=[(1, '1'), (2, '2'), (3, '3'), (4, '4'), (5, '5'), (6, '6'), (7, '7'), (8, '8')], default=1)),
                ('new_students', models.IntegerField(default=0)),
                ('chetlashtirilgan_students', models.IntegerField(default=0)),
                ('akademik', models.IntegerField(default=0)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('bosqich', models.CharField(choices=[('Bakalavr', 'Bakalavr'), ('Magistr', 'Magistr'), ('Aspirantura', 'Aspirantura'), ('Doktorantura', 'Doktorantura')], default='Bakalavr', max_length=100)),
            ],
            options={
                'verbose_name_plural': 'Guruhlar',
            },
        ),
        migrations.CreateModel(
            name='Budjet',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('jami', models.IntegerField(default=0)),
                ('harbiy', models.IntegerField(default=0)),
                ('xotin_qiz', models.IntegerField(default=0)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('guruhi', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.guruh')),
            ],
            options={
                'verbose_name_plural': 'Budjetlar',
            },
        ),
        migrations.CreateModel(
            name='Organization',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('img', models.ImageField(blank=True, null=True, upload_to='organization/')),
                ('user', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to=settings.AUTH_USER_MODEL)),
            ],
        ),
        migrations.AddField(
            model_name='guruh',
            name='org',
            field=models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.organization'),
        ),
        migrations.CreateModel(
            name='Faculty',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('org', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.organization')),
            ],
            options={
                'verbose_name_plural': 'Fakultetlar',
            },
        ),
        migrations.CreateModel(
            name='Shartnoma',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('jami', models.IntegerField(default=0)),
                ('harbiy', models.IntegerField(default=0)),
                ('xotin_qiz', models.IntegerField(default=0)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('guruh', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.guruh')),
            ],
            options={
                'verbose_name_plural': 'Shartnomalar',
            },
        ),
        migrations.CreateModel(
            name='UserOrg',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('org', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to='user.organization')),
                ('user', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'verbose_name_plural': 'Tashkilotlar',
            },
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
        migrations.CreateModel(
            name='Yonalish',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('code', models.CharField(blank=True, max_length=100, null=True)),
                ('turi', models.CharField(choices=[('Kunduzgi', 'Kunduzgi'), ('Sirtqi', 'Sirtqi'), ('Masofaviy', 'Masofaviy')], default='Kunduzgi', max_length=100)),
                ('language', models.CharField(choices=[("O'zbek", "O'zbek"), ('Rus', 'Rus')], default="O'zbek", max_length=100)),
                ('mutahasislik_2', models.BooleanField(default=False)),
                ('faculty', models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='facultyyonalish', to='user.faculty')),
                ('org', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.organization')),
            ],
            options={
                'verbose_name_plural': 'Yonalishlar',
            },
        ),
        migrations.AddField(
            model_name='guruh',
            name='yonalish',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name='yonalishguruh', to='user.yonalish'),
        ),
    ]
