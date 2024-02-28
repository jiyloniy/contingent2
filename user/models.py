from django.db import models
from guardian.shortcuts import assign_perm
from django.contrib.auth.models import User, Permission
from guardian.models import UserObjectPermission, GroupObjectPermission
from django.contrib.contenttypes.models import ContentType
from django.contrib.auth.models import Group
from django.db.models.signals import post_save
from django.dispatch import receiver


class   UserRules(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    org = models.ForeignKey('Organization', on_delete=models.CASCADE, null=True, blank=True)
    movqesi = models.CharField(max_length=100)
    password = models.CharField(max_length=100, null=True, blank=True)
    username = models.CharField(max_length=100, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    can_add_faculty = models.BooleanField(default=False)
    can_add_yonalish = models.BooleanField(default=False)
    can_add_guruh = models.BooleanField(default=False)
    can_update_faculty = models.BooleanField(default=False)
    can_update_yonalish = models.BooleanField(default=False)
    can_update_guruh = models.BooleanField(default=False)
    can_delete_faculty = models.BooleanField(default=False)
    can_delete_yonalish = models.BooleanField(default=False)
    can_delete_guruh = models.BooleanField(default=False)
    can_view_faculty = models.BooleanField(default=False)
    can_view_yonalish = models.BooleanField(default=False)
    can_view_guruh = models.BooleanField(default=   False)
    full_access = models.BooleanField(default=False)
    can_delete_user = models.BooleanField(default=False)
    can_add_user = models.BooleanField(default=False)
    can_update_user = models.BooleanField(default=False)
    can_view_user = models.BooleanField(default=False)
    def assign_perm(self, user, perm_codename):
        content_type = ContentType.objects.get_for_model(self.__class__)
        permission = Permission.objects.get(
            codename=perm_codename,
            content_type=content_type,
        )
        user.user_permissions.add(permission)

    def remove_perm(self, user, perm_codename):
        content_type = ContentType.objects.get_for_model(self.__class__)
        permission = Permission.objects.get(
            codename=perm_codename,
            content_type=content_type,
        )
        user.user_permissions.remove(permission)

    def __str__(self):
        return self.user.username

    class Meta:
        verbose_name_plural = 'Foydalanuvchilar'


class Organization(models.Model):
    user = models.ForeignKey(User, models.SET_NULL, blank=True, null=True)
    name = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)
    img = models.ImageField(upload_to='organization/', null=True, blank=True)

    def __str__(self):
        return self.name


class UserOrg(models.Model):
    org = models.ForeignKey(Organization, models.SET_NULL, blank=True, null=True)
    user = models.ForeignKey(User, models.SET_NULL, blank=True, null=True)

    def assign_perm(self, user, perm_codename):
        content_type = ContentType.objects.get_for_model(self.__class__)
        permission = Permission.objects.get(
            codename=perm_codename,
            content_type=content_type,
        )
        user.user_permissions.add(permission)

    def remove_perm(self, user, perm_codename):
        content_type = ContentType.objects.get_for_model(self.__class__)
        permission = Permission.objects.get(
            codename=perm_codename,
            content_type=content_type,
        )
        user.user_permissions.remove(permission)

    def __str__(self):
        return self.user.username

    class Meta:
        verbose_name_plural = 'Tashkilotlar'


class Faculty(models.Model):
    org = models.ForeignKey(Organization, on_delete=models.CASCADE)

    name = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name_plural = 'Fakultetlar'


class Yonalish(models.Model):
    org = models.ForeignKey(Organization, on_delete=models.CASCADE)
    name = models.CharField(max_length=100)
    faculty = models.ForeignKey(Faculty, on_delete=models.SET_NULL, blank=True, null=True,
                                related_name='facultyyonalish')
    created_at = models.DateTimeField(auto_now_add=True)

    turi_choices = (
        ('Kunduzgi', 'Kunduzgi'),
        ('Sirtqi', 'Sirtqi'),
        ('Masofaviy', 'Masofaviy'),
    )

    turi = models.CharField(max_length=100, choices=turi_choices, default='Kunduzgi')

    choises_uzb_rus = (
        ('O\'zbek', 'O\'zbek'),
        ('Rus', 'Rus'),
    )
    language = models.CharField(max_length=100, choices=choises_uzb_rus, default='O\'zbek')

    def __str__(self):
        return self.name

    class Meta:
        verbose_name_plural = 'Yonalishlar'


class Shartnoma(models.Model):
    guruh = models.ForeignKey('Guruh', on_delete=models.CASCADE)
    jami = models.IntegerField(default=0)
    harbiy = models.IntegerField(default=0)
    xotin_qiz = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'{self.jami} {self.harbiy} {self.xotin_qiz}'

    class Meta:
        verbose_name_plural = 'Shartnomalar'


class Budjet(models.Model):
    guruhi = models.ForeignKey('Guruh', on_delete=models.CASCADE)
    jami = models.IntegerField(default=0)
    harbiy = models.IntegerField(default=0)
    xotin_qiz = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'{self.jami} {self.harbiy} {self.xotin_qiz}'

    class Meta:
        verbose_name_plural = 'Budjetlar'


class Guruh(models.Model):
    org = models.ForeignKey(Organization, on_delete=models.CASCADE)
    name = models.CharField(max_length=100)
    yonalish = models.ForeignKey(Yonalish, on_delete=models.SET_NULL, blank=True, null=True,
                                 related_name='yonalishguruh')
    kurs_choices = (
        (1, '1'),
        (2, '2'),
        (3, '3'),
        (4, '4'),
        (5, '5'),
        (6, '6'),
        (7, '7'),
        (8, '8'),
    )

    kurs = models.IntegerField(choices=kurs_choices, default=1)

    new_students = models.IntegerField(default=0)
    chetlashtirilgan_students = models.IntegerField(default=0)
    akademik = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    type_choices = (
        ('Bakalavr', 'Bakalavr'),
        ('Magistr', 'Magistr'),
        ('Aspirantura', 'Aspirantura'),
        ('Doktorantura', 'Doktorantura'),
    )
    bosqich = models.CharField(max_length=100, choices=type_choices, default='Bakalavr')

    def __str__(self):
        return self.name

    class Meta:
        verbose_name_plural = 'Guruhlar'
