from django import forms
from django.forms import inlineformset_factory
from django.contrib.auth.models import User
from django.http import request

from user.models import Guruh, Budjet, Shartnoma, Faculty, Yonalish, UserOrg, Organization, UserRules


class UserOrgForm(forms.ModelForm):
    username = forms.CharField(max_length=100, label='Login')
    password = forms.CharField(max_length=100, label='Parol', widget=forms.PasswordInput)

    widgets = {
        'username': forms.TextInput(attrs={'class': 'form-control'}),
        'password': forms.PasswordInput(attrs={'class': 'form-control'})
    }

    labels = {
        'username': 'Login',
        '           password': 'Parol',
    }

    def clean(self):
        username = self.cleaned_data.get('username')
        password = self.cleaned_data.get('password')
        if not UserOrg.objects.filter(username=username).exists():
            raise forms.ValidationError('Username or password is incorrect')
        return self.cleaned_data


class GuruhForm(forms.ModelForm):
    class Meta:
        model = Guruh
        fields = ['name', 'yonalish', 'kurs', 'new_students', 'chetlashtirilgan_students', 'akademik', 'bosqich']

        labels = {
            'name': 'Guruh nomi',
            'yonalish': 'Yonalish',
            'kurs': 'Kurs',
            'new_students': 'Yangi studentlar',
            'chetlashtirilgan_students': 'Chetlashtirilgan studentlar',
            'akademik': 'Akademik',
            'bosqich': 'Bosqich'
        }

        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'yonalish': forms.Select(attrs={'class': 'form-control'}),
            'kurs': forms.Select(attrs={'class': 'form-control'}),
            'bosqich': forms.Select(attrs={'class': 'form-control'}),
            'new_students': forms.NumberInput(attrs={'class': 'form-control'}),
            'chetlashtirilgan_students': forms.NumberInput(attrs={'class': 'form-control'}),
            'akademik': forms.NumberInput(attrs={'class': 'form-control'})
        }

    # def __init__(self, *args, **kwargs):
    #     user = kwargs.pop('user', None)
    #     super(YonalishForm, self).__init__(*args, **kwargs)
    #     if user:
    #         if UserOrg.objects.filter(user=user).exists():
    #             self.fields['faculty'].queryset = Faculty.objects.filter(org=UserOrg.objects.get(user=user).org)
    #         if Organization.objects.filter(user=user).exists():
    #             self.fields['faculty'].queryset = Faculty.objects.filter(org=Organization.objects.get(user=user))
    def __init__(self, *args, **kwargs):
        user = kwargs.pop('user', None)
        super(GuruhForm, self).__init__(*args, **kwargs)

        if user:
            if UserOrg.objects.filter(user=user).exists():
                org = UserOrg.objects.get(user=user).org
                self.fields['yonalish'].queryset = Yonalish.objects.filter(org=org)
            elif Organization.objects.filter(user=user).exists():
                org = Organization.objects.get(user=user)
                self.fields['yonalish'].queryset = Yonalish.objects.filter(org=org)
        else:
            self.fields['yonalish'].queryset = Yonalish.objects.none()


class BudjetForm(forms.ModelForm):
    class Meta:
        model = Budjet
        fields = ['jami', 'harbiy', 'xotin_qiz']

        labels = {
            'jami': 'Jami',
            'harbiy': 'Harbiy',
            'xotin_qiz': 'Xotn qizlar'
        }
        widgets = {
            'jami': forms.NumberInput(attrs={'class': 'form-control'}),
            'harbiy': forms.NumberInput(attrs={'class': 'form-control'}),
            'xotin_qiz': forms.NumberInput(attrs={'class': 'form-control'})
        }


class ShartnomaForm(forms.ModelForm):
    class Meta:
        model = Shartnoma
        fields = ['jami', 'harbiy', 'xotin_qiz']
        labels = {
            'jami': 'Jami',
            'harbiy': 'Harbiy',
            'xotin_qiz': 'Xotn qizlar'
        }
        widgets = {
            'jami': forms.NumberInput(attrs={'class': 'form-control'}),
            'harbiy': forms.NumberInput(attrs={'class': 'form-control'}),
            'xotin_qiz': forms.NumberInput(attrs={'class': 'form-control'})
        }


ShartmonomaFormSet = inlineformset_factory(Guruh, Shartnoma, form=ShartnomaForm, extra=1, can_delete=False, max_num=1)
BudjetFormSet = inlineformset_factory(Guruh, Budjet, form=BudjetForm, extra=1, can_delete=False, max_num=1)


class FacultyForm(forms.ModelForm):
    class Meta:
        model = Faculty
        fields = ['name']

        labels = {
            'name': 'Nomi'
        }

        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'})
        }


#
# class YonalishForm(forms.ModelForm):
#     class Meta:
#         model = Yonalish
#         fields = ['name', 'faculty', 'turi', 'language']
#
#         labels = {
#             'name': 'Nomi',
#             'faculty': 'Fakultet',
#             'turi': 'Turi',
#             'language': 'Tili'
#         }
#         widgets = {
#             'name': forms.TextInput(attrs={'class': 'form-control'}),
#             'faculty': forms.Select(attrs={'class': 'form-control'}),
#             'turi': forms.Select(attrs={'class': 'form-control'}),
#             'language': forms.Select(attrs={'class': 'form-control'})
#         }
#
#     def __init__(self, *args, **kwargs):
#         user = kwargs.pop('user', None)
#         super(YonalishForm, self).__init__(*args, **kwargs)
#         print(user)
#         if user:
#             print(user)
#             if UserOrg.objects.filter(user=user).exists():
#                 org = UserOrg.objects.get(user=user).org
#                 self.fields['faculty'].queryset = Faculty.objects.filter(org=org)
#                 print(Faculty.objects.filter(org=UserOrg.objects.get(user=user).org))
#             if Organization.objects.filter(user=user).exists():
#                 print(Organization.objects.get(user=user))
#                 org = Organization.objects.get(user=user)
#                 self.fields['faculty'].queryset = Faculty.objects.filter(org=org)


class YonalishForm(forms.ModelForm):
    class Meta:
        model = Yonalish
        fields = ['name', 'faculty', 'turi', 'language','code','mutahasislik_2']

        labels = {
            'name': 'Nomi',
            'code':'Ko\'di',
            'faculty': 'Fakultet',
            'turi': 'Turi',
            'language': 'Tili',
            'mutahasislik_2':'2-mutahasislik'
        }

        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'code': forms.TextInput(attrs={'class': 'form-control'}),
            'faculty': forms.Select(attrs={'class': 'form-control'}),
            'turi': forms.Select(attrs={'class': 'form-control'}),
            'language': forms.Select(attrs={'class': 'form-control'}),
            'mutahasislik_2': forms.CheckboxInput(attrs={'class': 'form-check-input'})

        }

    def __init__(self, *args, **kwargs):
        user = kwargs.pop('user', None)
        print(user)
        super(YonalishForm, self).__init__(*args, **kwargs)

        if user:
            if UserOrg.objects.filter(user=user).exists():
                org = UserOrg.objects.get(user=user).org
                self.fields['faculty'].queryset = Faculty.objects.filter(org=org)

            elif Organization.objects.filter(user=user).exists():
                org = Organization.objects.get(user=user)
                self.fields['faculty'].queryset = Faculty.objects.filter(org=org)


class LoginForm(forms.Form):
    username = forms.CharField(max_length=100, label='Login')
    password = forms.CharField(max_length=100, label='Parol', widget=forms.PasswordInput)

    widgets = {
        'username': forms.TextInput(attrs={'class': 'form-control'}),
        'password': forms.PasswordInput(attrs={'class': 'form-control'})
    }

    labels = {
        'username': 'Login',
        'password': 'Parol',
    }

    def clean(self):
        username = self.cleaned_data.get('username')
        password = self.cleaned_data.get('password')
        user = User.objects.filter(username=username).first()
        if not user:
            raise forms.ValidationError('Username or password is incorrect')
        return self.cleaned_data


# Compare this snippet from user/views.py:

# UserOrg creation fomrs
# class UserRules(forms.Form):
#     username = forms.CharField(max_length=100, label='Login')
#     password = forms.CharField(max_length=100, label='Parol', widget=forms.PasswordInput)
#     # rules
#     can_add_faculty = forms.BooleanField(label='Fakultet qo`shish', required=False)
#     can_update_faculty = forms.BooleanField(label='Fakultetni tahrirlash', required=False)
#     can_delete_faculty = forms.BooleanField(label='Fakultetni o`chirish', required=False)
#     can_view_faculty = forms.BooleanField(label='Fakultetni ko`rish', required=False)
#     can_add_yonalish = forms.BooleanField(label='Yonalish qo`shish', required=False)
#     can_update_yonalish = forms.BooleanField(label='Yonalishni tahrirlash', required=False)
#     can_delete_yonalish = forms.BooleanField(label='Yonalishni o\'chirish', required=False)
#     can_view_yonalish = forms.BooleanField(label='Yonalishni ko\'rish', required=False)
#     can_add_guruh = forms.BooleanField(label='Guruh qo\'shish', required=False)
#     can_update_guruh = forms.BooleanField(label='Guruhni tahrirlash', required=False)
#     can_delete_guruh = forms.BooleanField(label='Guruhni o\'chirish', required=False)
#     can_view_guruh = forms.BooleanField(label='Guruhni ko\'rish', required=False)
#     can_add_user = forms.BooleanField(label='Foydalanuvchi qo\'shish', required=False)
#     can_update_user = forms.BooleanField(label='Foydalanuvchini tahrirlash', required=False)
#     can_delete_user = forms.BooleanField(label='Foydalanuvchini o\'chirish', required=False)
#     can_view_user = forms.BooleanField(label='Foydalanuvchini ko\'rish', required=False)
#     full_access = forms.BooleanField(label='Barcha huquqlar', required=False)
#
#     widgets = {
#         'username': forms.TextInput(attrs={'class': 'form-control'}),
#         'password': forms.PasswordInput(attrs={'class': 'form-control'}),
#         'can_add_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_update_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_delete_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_view_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_add_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_update_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_delete_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_view_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_add_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_update_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_delete_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_view_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_add_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_update_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'can_delete_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#         'full_access': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
#
#     }
#
#     labels = {
#         'username': 'Login',
#         'password': 'Parol',
#         'can_add_faculty': 'Fakultet qo`shish',
#         'can_update_faculty': 'Fakultetni tahrirlash',
#         'can_delete_faculty': 'Fakultetni o`chirish',
#         'can_view_faculty': 'Fakultetni ko`rish',
#         'can_add_yonalish': 'Yonalish qo`shish',
#         'can_update_yonalish': 'Yonalishni tahrirlash',
#         'can_delete_yonalish': 'Yonalishni o\'chirish',
#         'can_view_yonalish': 'Yonalar ko\'rish',
#         'can_add_guruh': 'Guruh qo\'shish',
#         'can_update_guruh': 'Guruhni tahrirlash',
#         'can_delete_guruh': 'Guruhni o\'chirish',
#         'can_view_guruh': 'Guruhni ko\'rish',
#         'can_add_user': 'Foydalanuvchi qo\'shish',
#         'can_update_user': 'Foydalanuvchini tahrirlash',
#         'can_delete_user': 'Foydalanuvchini o\'chirish',
#         'full_access': 'Barcha huquqlar',
#
#     }
#
#     class Meta:
#         model = UserOrg
#


class UserRule(forms.ModelForm):
    class Meta:
        model = UserRules
        fields = ['username', 'password', 'movqesi', 'full_access', 'can_add_faculty', 'can_update_faculty',
                  'can_delete_faculty',
                  'can_view_faculty', 'can_add_yonalish', 'can_update_yonalish', 'can_delete_yonalish',
                  'can_view_yonalish', 'can_add_guruh', 'can_update_guruh', 'can_delete_guruh', 'can_view_guruh',
                  'can_add_user', 'can_update_user', 'can_delete_user', 'can_view_user']

        widgets = {
            'username': forms.TextInput(attrs={'class': 'form-control'}),
            'password': forms.TextInput(attrs={'class': 'form-control'}),
            'movqesi': forms.TextInput(attrs={'class': 'form-control'}),
            'can_add_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_update_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_delete_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_view_faculty': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_add_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_update_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_delete_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_view_yonalish': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_add_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_update_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_delete_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_view_guruh': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_add_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_update_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_delete_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'full_access': forms.CheckboxInput(attrs={'class': 'form-check-input'}),
            'can_view_user': forms.CheckboxInput(attrs={'class': 'form-check-input'}),

        }

        labels = {
            'username': 'Login',
            'password': 'Parol',
            'movqesi': 'Mavqesi',
            'can_add_faculty': 'Fakultet qo`shish',
            'can_update_faculty': 'Fakultetni tahrirlash',
            'can_delete_faculty': 'Fakultetni o`chirish',
            'can_view_faculty': 'Fakultetni ko`rish',
            'can_add_yonalish': 'Yonalish qo`shish',
            'can_update_yonalish': 'Yonalishni tahrirlash',
            'can_delete_yonalish': 'Yonalishni o\'chirish',
            'can_view_yonalish': 'Yonalar ko\'rish',
            'can_add_guruh': 'Guruh qo\'shish',
            'can_update_guruh': 'Guruhni tahrirlash',
            'can_delete_guruh': 'Guruhni o\'chirish',
            'can_view_guruh': 'Guruhni ko\'rish',
            'can_add_user': 'Foydalanuvchi qo\'shish',
            'can_update_user': 'Foydalanuvchini tahrirlash',
            'can_delete_user': 'Foydalanuvchini o\'chirish',
            'full_access': 'Barcha huquqlar',
            'can_view_user': 'Foydalanuvchini ko\'rish',
        }



class UserRuleUpdate(forms.ModelForm):

    class Meta:
        model = UserRules
        fields = '__all__'
        exclude = ['user', 'org']





