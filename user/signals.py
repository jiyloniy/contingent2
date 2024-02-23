from user.models import UserProfile, Kurs, Faculty, Yonalish, Organization, Shartnoma, Budjet, Guruh
from django.db.models.signals import post_save, pre_save, pre_delete, post_delete, m2m_changed, pre_init, post_init
from django.dispatch import receiver
from django.contrib.auth.models import User
from django.db.models import Q
from guardian.shortcuts import assign_perm, remove_perm



# full rule list

def full_role(sender, instance, created, **kwargs):
    if created:
        UserProfile.objects.create(user=instance, organization=instance)
        assign_perm('change_userprofile', instance, instance.userprofile)
        assign_perm('delete_userprofile', instance, instance.userprofile)
        assign_perm('view_userprofile', instance, instance.userprofile)
        assign_perm('add_userprofile', instance, instance.userprofile)
        assign_perm('change_kurs', instance, instance.kurs)
        assign_perm('delete_kurs', instance, instance.kurs)
        assign_perm('view_kurs', instance, instance.kurs)
        assign_perm('add_kurs', instance, instance.kurs)
        assign_perm('change_faculty', instance, instance.faculty)
        assign_perm('delete_faculty', instance, instance.faculty)
        assign_perm('view_faculty', instance, instance.faculty)
        assign_perm('add_faculty', instance, instance.faculty)
        assign_perm('change_yonalish', instance, instance.yonalish)
        assign_perm('delete_yonalish', instance, instance.yonalish)
        assign_perm('view_yonalish', instance, instance.yonalish)
        assign_perm('add_yonalish', instance, instance.yonalish)
        assign_perm('change_shartnoma', instance, instance.shartnoma)
        assign_perm('delete_shartnoma', instance, instance.shartnoma)
        assign_perm('view_shartnoma', instance, instance.shartnoma)
        assign_perm('add_shartnoma', instance, instance.shartnoma)
        assign_perm('change_budjet', instance, instance.budjet)
        assign_perm('delete_budjet', instance, instance.budjet)
        assign_perm('view_budjet', instance, instance.budjet)
        assign_perm('add_budjet', instance, instance.budjet)
        assign_perm('change_guruh', instance, instance.guruh)
        assign_perm('delete_guruh', instance, instance.guruh)
        assign_perm('view_guruh', instance, instance.guruh)
        assign_perm('add_guruh', instance, instance.guruh)


# rule for can change kurs
def fullkurs(sender, instance, **kwargs):
    assign_perm('change_kurs', instance.org.user, instance)
    assign_perm('delete_kurs', instance.org.user, instance)
    assign_perm('view_kurs', instance.org.user, instance)
    assign_perm('add_kurs', instance.org.user, instance)


def changekurs(sender, instance, **kwargs):
    assign_perm('change_kurs', instance.org.user, instance)


def deletekurs(sender, instance, **kwargs):
    assign_perm('delete_kurs', instance.org.user, instance)


def viewkurs(sender, instance, **kwargs):
    assign_perm('view_kurs', instance.org.user, instance)


def addkurs(sender, instance, **kwargs):
    assign_perm('add_kurs', instance.org.user, instance)


def fullfaculty(sender, instance, **kwargs):
    assign_perm('change_faculty', instance.org.user, instance)
    assign_perm('delete_faculty', instance.org.user, instance)
    assign_perm('view_faculty', instance.org.user, instance)
    assign_perm('add_faculty', instance.org.user, instance)


def changefaculty(sender, instance, **kwargs):
    assign_perm('change_faculty', instance.org.user, instance)


def deletefaculty(sender, instance, **kwargs):
    assign_perm('delete_faculty', instance.org.user, instance)


def viewfaculty(sender, instance, **kwargs):
    assign_perm('view_faculty', instance.org.user, instance)


def addfaculty(sender, instance, **kwargs):
    assign_perm('add_faculty', instance.org.user, instance)


def fullyonalish(sender, instance, **kwargs):
    assign_perm('change_yonalish', instance.org.user, instance)
    assign_perm('delete_yonalish', instance.org.user, instance)
    assign_perm('view_yonalish', instance.org.user, instance)
    assign_perm('add_yonalish', instance.org.user, instance)


def changeyonalish(sender, instance, **kwargs):
    assign_perm('change_yonalish', instance.org.user, instance)


def deleteyonalish(sender, instance, **kwargs):
    assign_perm('delete_yonalish', instance.org.user, instance)


def viewyonalish(sender, instance, **kwargs):
    assign_perm('view_yonalish', instance.org.user, instance)


def addyonalish(sender, instance, **kwargs):
    assign_perm('add_yonalish', instance.org.user, instance)


def fullguhruh(sender, instance, **kwargs):
    assign_perm('change_guruh', instance.org.user, instance)
    assign_perm('delete_guruh', instance.org.user, instance)
    assign_perm('view_guruh', instance.org.user, instance)
    assign_perm('add_guruh', instance.org.user, instance)
    assign_perm('change_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('delete_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('view_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('add_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('change_budjet', instance.org.user, instance.budjet)
    assign_perm('delete_budjet', instance.org.user, instance.budjet)
    assign_perm('view_budjet', instance.org.user, instance.budjet)
    assign_perm('add_budjet', instance.org.user, instance.budjet)


def changeguruh(sender, instance, **kwargs):
    assign_perm('change_guruh', instance.org.user, instance)
    assign_perm('change_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('change_budjet', instance.org.user, instance.budjet)


def deleteguruh(sender, instance, **kwargs):
    assign_perm('delete_guruh', instance.org.user, instance)
    assign_perm('delete_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('delete_budjet', instance.org.user, instance.budjet)


def viewguruh(sender, instance, **kwargs):
    assign_perm('view_guruh', instance.org.user, instance)
    assign_perm('view_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('view_budjet', instance.org.user, instance.budjet)


def addguruh(sender, instance, **kwargs):
    assign_perm('add_guruh', instance.org.user, instance)
    assign_perm('add_shartnoma', instance.org.user, instance.shartnoma)
    assign_perm('add_budjet', instance.org.user, instance.budjet)
