# Generated by Django 4.1.7 on 2023-09-26 04:35

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Register',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('Name', models.CharField(max_length=60)),
                ('E_mail', models.CharField(max_length=50)),
                ('password', models.CharField(max_length=50)),
                ('Re_password', models.CharField(max_length=50)),
            ],
        ),
    ]
