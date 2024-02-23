# Generated by Django 5.0 on 2024-02-21 16:24

import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('user', '0003_userrulser'),
    ]

    operations = [
        migrations.CreateModel(
            name='Faculty',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('org', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.organization')),
                ('user', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.userrulser')),
            ],
            options={
                'verbose_name_plural': 'Fakultetlar',
            },
        ),
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
                ('org', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='user.organization')),
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
            name='Yonalish',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=100)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('turi', models.CharField(choices=[('Kunduzgi', 'Kunduzgi'), ('Sirtqi', 'Sirtqi'), ('Masofaviy', 'Masofaviy')], default='Kunduzgi', max_length=100)),
                ('language', models.CharField(choices=[("O'zbek", "O'zbek"), ('Rus', 'Rus')], default="O'zbek", max_length=100)),
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