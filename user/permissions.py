from django.core.exceptions import ValidationError

from user.forms import UserRule


class UserPermissions:

    def __init__(self, user):
        self.user = user

    # if form.cleaned_data.get('full_access'):
    #     user.user_permissions.add(Permission.objects.get(codename='view_user'))
    #     user.user_permissions.add(Permission.objects.get(codename='add_user'))
    #     user.user_permissions.add(Permission.objects.get(codename='change_user'))
    #     user.user_permissions.add(Permission.objects.get(codename='delete_user'))
    #     user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
    #     user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
    #     user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
    #     user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
    #     user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
    #     user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
    #     user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
    #     user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
    #     user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
    #     user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
    #     user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
    #     user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))

    def has_full_access(self):
        return self.user.has_perm('user.view_user') and self.user.has_perm('user.add_user') and self.user.has_perm(
            'user.change_user') and self.user.has_perm('user.delete_user') and self.user.has_perm(
            'user.view_faculty') and self.user.has_perm('user.add_faculty') and self.user.has_perm(
            'user.change_faculty') and self.user.has_perm('user.delete_faculty') and self.user.has_perm(
            'user.view_yonalish') and self.user.has_perm('user.add_yonalish') and self.user.has_perm(
            'user.change_yonalish') and self.user.has_perm('user.delete_yonalish') and self.user.has_perm(
            'user.view_guruh') and self.user.has_perm('user.add_guruh') and self.user.has_perm(
            'user.change_guruh') and self.user.has_perm('user.delete_guruh') and self.user.has_perm(
            'user.view_userorg') and self.user.has_perm('user.add_userorg') and self.user.has_perm(
            'user.change_userorg') and self.user.has_perm('user.delete_userorg')

    def get_permissions(self):
        permissions = []
        if self.has_full_access():
            permissions.extend(['view_user', 'add_user', 'change_user', 'delete_user', 'view_faculty', 'add_faculty',
                                'change_faculty', 'delete_faculty', 'view_yonalish', 'add_yonalish', 'change_yonalish',
                                'delete_yonalish', 'view_guruh', 'add_guruh', 'change_guruh', 'delete_guruh',
                                'view_userorg', 'add_userorg', 'change_userorg', 'delete_userorg'])

        if self.user.has_perm('view_user'):
            permissions.append('view_user')

        if self.user.has_perm('add_user'):
            permissions.append('add_user')

        if self.user.has_perm('change_user'):
            permissions.append('change_user')

        if self.user.has_perm('delete_user'):
            permissions.append('delete_user')

        if self.user.has_perm('view_faculty'):
            permissions.append('view_faculty')

        if self.user.has_perm('add_faculty'):
            permissions.append('add_faculty')

        if self.user.has_perm('change_faculty'):
            permissions.append('change_faculty')

        if self.user.has_perm('delete_faculty'):
            permissions.append('delete_faculty')

        if self.user.has_perm('view_yonalish'):
            permissions.append('view_yonalish')

        if self.user.has_perm('add_yonalish'):
            permissions.append('add_yonalish')

        if self.user.has_perm('change_yonalish'):
            permissions.append('change_yonalish')

        if self.user.has_perm('delete_yonalish'):
            permissions.append('delete_yonalish')

        if self.user.has_perm('view_guruh'):
            permissions.append('view_guruh')

        if self.user.has_perm('add_guruh'):
            permissions.append('add_guruh')

        if self.user.has_perm('change_guruh'):
            permissions.append('change_guruh')

        if self.user.has_perm('delete_guruh'):
            permissions.append('delete_guruh')

        return permissions


def validate_form(request):
  form = UserRule(request.POST)
  if form.is_valid():
    return form
  else:
      raise ValidationError('Invalid form')


def create_user(form, request):
  username = form.cleaned_data['username']
