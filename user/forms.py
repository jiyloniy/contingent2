from django import forms
from django.forms import inlineformset_factory
from django.contrib.auth.models import User
from user.models import Guruh, Budjet, Shartnoma, Faculty, Yonalish, UserOrg, Organization


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
                self.fields['yonalish'].queryset = Yonalish.objects.filter(org=UserOrg.objects.get(user=user).org)
            if Organization.objects.filter(user=user).exists():
                self.fields['yonalish'].queryset = Yonalish.objects.filter(org=Organization.objects.get(user=user))


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
        fields = ['name', 'faculty', 'turi', 'language']

        labels = {
            'name': 'Nomi',
            'faculty': 'Fakultet',
            'turi': 'Turi',
            'language': 'Tili'
        }

        widgets = {
            'name': forms.TextInput(attrs={'class': 'form-control'}),
            'faculty': forms.Select(attrs={'class': 'form-control'}),
            'turi': forms.Select(attrs={'class': 'form-control'}),
            'language': forms.Select(attrs={'class': 'form-control'})
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
