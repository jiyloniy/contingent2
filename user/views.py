from django.db import transaction
from django.shortcuts import render
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout
from django.shortcuts import redirect
from django.contrib import messages
from django.contrib.auth.models import Group
from django.contrib.auth.models import Permission
from django.contrib.contenttypes.models import ContentType
from guardian.shortcuts import assign_perm
from django.contrib.auth.models import Group
from django.contrib.auth.models import Permission
from guardian.models import UserObjectPermission, GroupObjectPermission
from django.contrib.contenttypes.models import ContentType
from user.models import Organization, UserOrg, Budjet, Guruh, Shartnoma, Faculty, Yonalish

from user.forms import LoginForm, FacultyForm, YonalishForm, ShartnomaForm, BudjetForm, GuruhForm


def login_view(request):
    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if user:
                login(request, user)
                return redirect('dashboard')
            else:
                messages.error(request, 'Username or password is incorrect')
    else:
        form = LoginForm()
    return render(request, 'pages/login.html', {'form': form})


def logout_view(request):
    logout(request)
    return redirect('login')


@login_required(login_url='login')
def faculty_list(request):
    form = FacultyForm()
    # can view
    if request.user.user_permissions.filter(codename='view_faculty').exists():
        print('can_view')
        if UserOrg.objects.filter(user=request.user).exists():
            org = UserOrg.objects.get(user=request.user).org
            faculties = Faculty.objects.filter(org=org)
            context = {'faculties': faculties, 'form': form}
            return render(request, 'pages/dashboard.html', context)
        if Organization.objects.filter(user=request.user).exists():
            org = Organization.objects.get(user=request.user)
            faculties = Faculty.objects.filter(org=org)
            context = {'faculties': faculties, 'form': form}
            return render(request, 'pages/dashboard.html', context)
    return render(request, 'pages/dashboard.html', {'form': form})


@login_required(login_url='login')
def add_faculty(request):
    if request.user.user_permissions.filter(codename='add_faculty').exists():

        if request.method == 'POST':
            form = FacultyForm(request.POST)
            if form.is_valid():
                name = form.cleaned_data.get('name')
                print(name)

                if UserOrg.objects.filter(user=request.user).exists():
                    print(1)
                    org = UserOrg.objects.get(user=request.user).org
                    faculty = Faculty(name=name, org=org)
                    print(org)
                    faculty.save()

                if Organization.objects.filter(user=request.user).exists():
                    print(2)
                    org = Organization.objects.get(user=request.user)
                    print(org)
                    faculty = Faculty(name=name, org=org)
                    faculty.save()

        else:

            return redirect('dashboard')
    messages.error(request, 'You do not have permission to add faculty')
    return redirect('dashboard')


@login_required(login_url='login')
def faculty_update(request, pk):
    if request.user.user_permissions.filter(codename='change_faculty').exists():
        faculty = Faculty.objects.get(id=pk)
        form = FacultyForm(instance=faculty)
        if request.method == 'POST':
            form = FacultyForm(request.POST, instance=faculty)
            if form.is_valid():
                form.save()
                return redirect('dashboard')
        context = {'form': form}
        return render(request, 'baseapp/faculty_form.html', context)
    messages.error(request, 'You do not have permission to update faculty')
    return redirect('dashboard')


@login_required(login_url='login')
def faculty_delete(request, pk):
    if request.user.user_permissions.filter(codename='delete_faculty').exists():
        faculty = Faculty.objects.get(id=pk)
        if request.method == 'POST':
            faculty.delete()
            return redirect('dashboard')
        context = {'faculty': faculty}
        return redirect('dashboard')
    messages.error(request, 'You do not have permission to delete faculty')
    return redirect('dashboard')


@login_required(login_url='login')
def yonalish_list(request):
    form = YonalishForm(user=request.user)
    if request.user.user_permissions.filter(codename='view_yonalish').exists():
        if UserOrg.objects.filter(user=request.user).exists():
            org = UserOrg.objects.get(user=request.user).org
            yonalish = Yonalish.objects.filter(org=org)
            context = {'yonalish': yonalish, 'form': form}
            return render(request, 'pages/settings.html', context)
        if Organization.objects.filter(user=request.user).exists():
            org = Organization.objects.get(user=request.user)
            yonalish = Yonalish.objects.filter(org=org)
            context = {'yonalishlar': yonalish, 'form': form}
            return render(request, 'pages/settings.html', context)

    return render(request, 'pages/settings.html', {'form': form})


@login_required(login_url='login')
def yonalish_create(request):
    print(request.user)
    if request.user.user_permissions.filter(codename='add_yonalish').exists():
        user = User.objects.filter(username=request.user.username).first()
        if request.method == 'POST':
            form = YonalishForm(request.POST, user=user)
            if form.is_valid():
                name = form.cleaned_data.get('name')
                faculty = form.cleaned_data.get('faculty')
                turi = form.cleaned_data.get('turi')
                language = form.cleaned_data.get('language')
                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org
                    yonalish = Yonalish(name=name, faculty=faculty, turi=turi, language=language, org=org)
                    yonalish.save()
                if Organization.objects.filter(user=request.user).exists():
                    org = Organization.objects.get(user=request.user)
                    yonalish = Yonalish(name=name, faculty=faculty, turi=turi, language=language, org=org)
                    yonalish.save()

        else:
            return redirect('yonalishlar')

    messages.error(request, 'You do not have permission to add yonalish')
    return redirect('yonalishlar')


@login_required(login_url='login')
def yonalish_update(request, pk):
    if request.user.user_permissions.filter(codename='change_yonalish').exists():
        yonalish = Yonalish.objects.get(id=pk)
        form = YonalishForm(instance=yonalish, user=request.user)
        if request.method == 'POST':
            form = YonalishForm(request.POST, instance=yonalish, user=request.user)
            if form.is_valid():
                form.save()
                return redirect('yonalishlar')
        context = {'form': form}
        return render(request, 'baseapp/yonalish_form.html', context)
    messages.error(request, 'You do not have permission to update yonalish')
    return redirect('yonalishlar')


@login_required(login_url='login')
def yonalish_delete(request, pk):
    if request.user.user_permissions.filter(codename='delete_yonalish').exists():
        yonalish = Yonalish.objects.get(id=pk)
        if request.method == 'POST':
            yonalish.delete()
            return redirect('yonalishlar')
        context = {'yonalish': yonalish}
        return redirect('yonalishlar')
    messages.error(request, 'You do not have permission to delete yonalish')
    return redirect('yonalishlar')


@login_required(login_url='login')
def guruh_list(request):
    form = GuruhForm()
    shartmonoma = ShartnomaForm()
    budjet = BudjetForm()
    if request.user.user_permissions.filter(codename='view_guruh').exists():
        if UserOrg.objects.filter(user=request.user).exists():
            org = UserOrg.objects.get(user=request.user).org
            guruh = Guruh.objects.filter(org=org)
            context = {'guruh': guruh, 'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet}
            return render(request, 'pages/tables.html', context)
        if Organization.objects.filter(user=request.user).exists():
            org = Organization.objects.get(user=request.user)
            guruh = Guruh.objects.filter(org=org)
            context = {'guruh': guruh, 'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet}
            return render(request, 'pages/tables.html', context)
        else:
            return render(request, 'pages/tables.html', {'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet})
    return render(request, 'pages/tables.html', {'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet})


@login_required(login_url='login')
def guruh_create(request):
    if request.user.user_permissions.filter(codename='add_guruh').exists():
        if request.method == 'POST':
            form = GuruhForm(request.POST, user=request.user)
            budjet = BudjetForm(request.POST)
            shartmonoma = ShartnomaForm(request.POST)
            if form.is_valid() and budjet.is_valid() and shartmonoma.is_valid():
                name = form.cleaned_data.get('name')
                yonalish = form.cleaned_data.get('yonalish')
                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org
                    with transaction.atomic():
                        guruh = Guruh(name=name, yonalish=yonalish, org=org)
                        guruh.save()
                        shartmonoma = Shartnoma(jami=shartmonoma.cleaned_data.get('jami'),
                                                harbiy=shartmonoma.cleaned_data.get('harbiy'),
                                                xotin_qiz=shartmonoma.cleaned_data.get('xotin_qiz'), guruh=guruh)
                        shartmonoma.save()
                        budjet = Budjet(jami=budjet.cleaned_data.get('jami'), harbiy=budjet.cleaned_data.get('harbiy'),
                                        xotin_qiz=budjet.cleaned_data.get('xotin_qiz'), guruh=guruh)
                        budjet.save()
                        return redirect('guruh')
                if Organization.objects.filter(user=request.user).exists():
                    org = Organization.objects.get(user=request.user)
                    with transaction.atomic():
                        guruh = Guruh(name=name, yonalish=yonalish, org=org)
                        guruh.save()
                        shartmonoma = Shartnoma(jami=shartmonoma.cleaned_data.get('jami'),
                                                harbiy=shartmonoma.cleaned_data.get('harbiy'),
                                                xotin_qiz=shartmonoma.cleaned_data.get('xotin_qiz'), guruh=guruh)
                        shartmonoma.save()
                        budjet = Budjet(jami=budjet.cleaned_data.get('jami'), harbiy=budjet.cleaned_data.get('harbiy'),
                                        xotin_qiz=budjet.cleaned_data.get('xotin_qiz'), guruh=guruh)
                        budjet.save()
                        return redirect('guruh')
        else:
            return redirect('guruhlar')
    messages.error(request, 'You do not have permission to add guruh')
    return redirect('guruh')


@login_required(login_url='login')
def guruh_update(request, pk):
    if request.user.user_permissions.filter(codename='change_guruh').exists():
        guruh = Guruh.objects.get(id=pk)
        form = GuruhForm(instance=guruh, user=request.user)
        shartmonoma = ShartnomaForm(instance=guruh.shartnoma)
        budjet = BudjetForm(instance=guruh.budjet)
        if request.method == 'POST':
            form = GuruhForm(request.POST, instance=guruh, user=request.user)
            shartmonoma = ShartnomaForm(request.POST, instance=guruh.shartnoma)
            budjet = BudjetForm(request.POST, instance=guruh.budjet)
            if form.is_valid() and shartmonoma.is_valid() and budjet.is_valid():
                form.save()
                shartmonoma.save()
                budjet.save()
                return redirect('guruh')
        context = {'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet}
        return render(request, 'baseapp/group_form.html', context)
    messages.error(request, 'You do not have permission to update guruh')
    return redirect('guruh')


@login_required(login_url='login')
def guruh_delete(request, pk):
    if request.user.user_permissions.filter(codename='delete_guruh').exists():
        guruh = Guruh.objects.get(id=pk)
        if request.method == 'POST':
            guruh.delete()
            return redirect('guruh')
        context = {'guruh': guruh}
        return redirect('guruh')
    messages.error(request, 'You do not have permission to delete guruh')
    return redirect('guruhlar')
