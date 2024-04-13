from django.db import models

# Create your models here.
class Register(models.Model):
    Name = models.CharField(max_length=60)
    E_mail = models.CharField(max_length=50)
    password = models.CharField(max_length=50)
    Re_password = models.CharField(max_length=50)

    def __str__(self):
        return self.Name