import os

from background_task import background
from background_task.models import Task
from django.contrib.sessions.models import Session
from django.db import transaction
from django.shortcuts import render
from django.http import HttpResponse
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.contrib.auth import authenticate, login, logout
from django.http import FileResponse
from django.shortcuts import redirect
from django.contrib import messages
from django.contrib.auth.models import Group
from django.contrib.auth.models import Permission
from django.contrib.contenttypes.models import ContentType

from excelgenerate import generate_excel_files, merge_excel_files

from django.contrib.auth.models import Group
from django.contrib.auth.models import Permission
from django.utils import timezone
from guardian.models import UserObjectPermission, GroupObjectPermission
from django.contrib.contenttypes.models import ContentType
from user.models import Guruh, Faculty, UserRules

from user.forms import LoginForm, FacultyForm, ShartnomaForm, BudjetForm, GuruhForm, ShartmonomaFormSet, \
    BudjetFormSet, UserRule, UserRuleUpdate


def login_view(request):
    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(username=username, password=password)
            if UserOrg.objects.filter(user=user).exists():
                if user:

                    # user_agent = request.META['HTTP_USER_AGENT']
                    # ip = request.META['REMOTE_ADDR']
                    # existing_session = Session.objects.filter(expire_date__gte=timezone.now())
                    # if existing_session.exists():
                    #     existing_session.delete()

                    login(request, user)
                    if request.user.user_permissions.filter(codename='view_faculty').exists():
                        return redirect('dashboard')
                    if request.user.user_permissions.filter(codename='view_yonalish').exists():
                        return redirect('yonalishlar')
                    if request.user.user_permissions.filter(codename='view_guruh').exists():
                        return redirect('guruh')
                    if request.user.user_permissions.filter(codename='view_user').exists():
                        return redirect('userlist')
                    return redirect('empty')

                else:
                    messages.error(request, 'Username or password is incorrect')
            if Organization.objects.filter(user=user).exists():
                if user:
                    login(request, user)
                    if request.user.user_permissions.filter(codename='view_faculty').exists():
                        return redirect('dashboard')
                    if request.user.user_permissions.filter(codename='view_yonalish').exists():
                        return redirect('yonalishlar')
                    if request.user.user_permissions.filter(codename='view_guruh').exists():
                        return redirect('guruh')
                    if request.user.user_permissions.filter(codename='view_user').exists():
                        return redirect('userlist')
                    return redirect('empty')
                else:
                    messages.error(request, 'Username or password is incorrect')
            else:
                messages.error(request, 'You do not have permission to login')
                redirect('login')
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
        else:
            messages.error(request, 'You do not have permission to view faculty')
            return redirect('empty')
    return redirect('empty')


@login_required(login_url='login')
def add_faculty(request):
    if request.user.user_permissions.filter(codename='add_faculty').exists():

        if request.method == 'POST':
            form = FacultyForm(request.POST)
            if form.is_valid():
                name = form.cleaned_data.get('name')

                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org
                    faculty = Faculty(name=name, org=org)

                    faculty.save()

                if Organization.objects.filter(user=request.user).exists():
                    org = Organization.objects.get(user=request.user)

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
        faculty.delete()
        return redirect('dashboard')
    else:
        messages.error(request, 'You do not have permission to delete faculty')
        return redirect('dashboard')


from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.shortcuts import render, redirect
from django.db.models import Prefetch

from .models import Yonalish, Organization, UserOrg
from .forms import YonalishForm


@login_required(login_url='login')
def yonalish_list(request):
    form = YonalishForm(user=request.user)

    if request.user.user_permissions.filter(codename='view_yonalish').exists():
        org = None

        # Use select_related to fetch related data in a single query
        user_org = UserOrg.objects.select_related('org').filter(user=request.user).first()
        if user_org:
            org = user_org.org
        else:
            org = Organization.objects.filter(user=request.user).first()

        if org:
            # Use prefetch_related to fetch related data in separate queries
            yonalish = Yonalish.objects.select_related('faculty').prefetch_related(
                Prefetch('org', queryset=Organization.objects.select_related('user'))
            ).filter(org=org)
            context = {'yonalishlar': yonalish, 'form': form}
            return render(request, 'pages/settings.html', context)
        else:
            messages.error(request, 'You do not have permission to view yonalish')
            return redirect('empty')
    else:
        return redirect('empty')


@login_required(login_url='login')
def yonalish_create(request):
    if request.user.user_permissions.filter(codename='add_yonalish').exists():
        user = User.objects.filter(username=request.user.username).first()
        if request.method == 'POST':
            form = YonalishForm(request.POST, user=user)
            if form.is_valid():
                name = form.cleaned_data.get('name')
                faculty = form.cleaned_data.get('faculty')
                turi = form.cleaned_data.get('turi')
                language = form.cleaned_data.get('language'),
                print(language[0], type(language))
                code = form.cleaned_data.get('code'),
                mutahasislik_2 = form.cleaned_data.get('mutahasislik_2')
                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org
                    yonalish = Yonalish(name=name, faculty=faculty, turi=turi, language=language[0], org=org, code=code,
                                        mutahasislik_2=mutahasislik_2)
                    yonalish.save()
                if Organization.objects.filter(user=request.user).exists():
                    org = Organization.objects.get(user=request.user)
                    yonalish = Yonalish(name=name, faculty=faculty, turi=turi, language=language, org=org, code=code,
                                        mutahasislik_2=mutahasislik_2)
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
        yonalish.delete()
        return redirect('yonalishlar')
    else:
        messages.error(request, 'You do not have permission to delete yonalish')
        return redirect('yonalishlar')


@login_required(login_url='login')
def guruh_list(request):
    form = GuruhForm()
    shartmonoma = ShartnomaForm()
    budjet = BudjetForm()

    if request.user.user_permissions.filter(codename='view_guruh').exists():
        org = None

        # Use select_related to fetch related data in a single query
        user_org = UserOrg.objects.select_related('org').filter(user=request.user).first()
        if user_org:
            org = user_org.org
        else:
            org = Organization.objects.filter(user=request.user).first()

        if org:
            # Use prefetch_related to fetch related data in separate queries
            guruh = Guruh.objects.select_related('yonalish').prefetch_related(
                Prefetch('org', queryset=Organization.objects.select_related('user'))
            ).filter(org=org)
            context = {'groups': guruh, 'form': form, 'shartmonoma': shartmonoma, 'budjet': budjet}
            return render(request, 'pages/tables.html', context)
        else:
            messages.error(request, 'Your organization information is missing.')
            return redirect('empty')
    else:
        messages.error(request, 'You do not have permission to view groups')
        return redirect('empty')

    return redirect('empty')


@login_required(login_url='login')
def guruh_create(request):
    form = GuruhForm(user=request.user)
    shartmonoma = ShartmonomaFormSet()
    budjet = BudjetFormSet()

    context = {
        'form': form,
        'shartmonoma': shartmonoma,
        'budjet': budjet
    }
    if request.user.user_permissions.filter(codename='add_guruh').exists():
        if request.method == 'POST':
            form = GuruhForm(request.POST, user=request.user)
            budjet = BudjetFormSet(request.POST)
            shartmonoma = ShartmonomaFormSet(request.POST)

            if form.is_valid() and budjet.is_valid() and shartmonoma.is_valid():
                name = form.cleaned_data.get('name')
                yonalish = form.cleaned_data.get('yonalish')
                kurs = form.cleaned_data.get('kurs')
                new_students = form.cleaned_data.get('new_students')
                chetlashtirilgan_students = form.cleaned_data.get('chetlashtirilgan_students')
                akademik = form.cleaned_data.get('akademik')
                bosqich = form.cleaned_data.get('bosqich')

                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org

                    with transaction.atomic():
                        guruh = Guruh(name=name, yonalish=yonalish, org=org, kurs=kurs, new_students=new_students,
                                      chetlashtirilgan_students=chetlashtirilgan_students, akademik=akademik,
                                      bosqich=bosqich)
                        guruh.save()
                        shartmonoma.save(commit=False)
                        shartmonoma.instance = guruh
                        shartmonoma.save()
                        budjet.save(commit=False)
                        budjet.instance = guruh
                        budjet.save()
                        return redirect('guruh')
                if Organization.objects.filter(user=request.user).exists():
                    org = Organization.objects.get(user=request.user)

                    with transaction.atomic():
                        guruh = Guruh(name=name, yonalish=yonalish, org=org, kurs=kurs, new_students=new_students,
                                      chetlashtirilgan_students=chetlashtirilgan_students, akademik=akademik,
                                      bosqich=bosqich)
                        guruh.save()
                        shartmonoma.save(commit=False)
                        shartmonoma.instance = guruh
                        shartmonoma.save()
                        return redirect('guruh')
        else:

            return render(request, 'baseapp/group_form.html', context)
    messages.error(request, 'You do not have permission to add guruh')
    return redirect('guruh')


@login_required(login_url='login')
def guruh_update(request, pk):
    if request.user.user_permissions.filter(codename='change_guruh').exists():
        guruh = Guruh.objects.get(id=pk)
        form = GuruhForm(instance=guruh, user=request.user)
        shartmonoma = ShartmonomaFormSet(instance=guruh)
        budjet = BudjetFormSet(instance=guruh)
        if request.method == 'POST':
            form = GuruhForm(request.POST, instance=guruh, user=request.user)
            shartmonoma = ShartmonomaFormSet(request.POST, instance=guruh)
            budjet = BudjetFormSet(request.POST, instance=guruh)
            if form.is_valid() and shartmonoma.is_valid() and budjet.is_valid():
                form.save()
                shartmonoma.save()
                budjet.save()
                return redirect('guruh')
        context = {
            'form': form,
            'shartmonoma': shartmonoma,
            'budjet': budjet
        }
        return render(request, 'baseapp/group_form.html', context)


@login_required(login_url='login')
def guruh_delete(request, pk):
    if request.user.user_permissions.filter(codename='delete_guruh').exists():
        guruh = Guruh.objects.get(id=pk)
        guruh.delete()
        return redirect('guruh')
    else:
        messages.error(request, 'You do not have permission to delete guruh')
        return redirect('guruh')


@login_required(login_url='login')
def userlist(request):
    if request.user.user_permissions.filter(codename='view_user').exists():
        org = None

        # Use select_related to fetch related data in a single query
        user_org = UserOrg.objects.select_related('user', 'org').filter(user=request.user).first()
        if user_org:
            org = user_org.org
        else:
            org = Organization.objects.filter(user=request.user).first()

        if org:
            # Use prefetch_related to fetch related data in separate queries
            users = UserOrg.objects.prefetch_related(
                Prefetch('user'),
                Prefetch('org', queryset=Organization.objects.select_related('user'))
            ).filter(org=org)
            context = {'users': users}
            return render(request, 'pages/usercreate.html', context)
        else:
            return render(request, 'pages/usercreate.html')
    else:
        messages.error(request, 'You do not have permission to view user')
        return redirect('empty')


@login_required(login_url='login')
def user_create(request):
    if request.user.user_permissions.filter(codename='add_user').exists():
        if request.method == 'POST':
            form = UserRule(request.POST)
            if form.is_valid():
                username = form.cleaned_data.get('username')
                password = form.cleaned_data.get('password')
                user = User.objects.create_user(username=username, password=password)
                if Organization.objects.filter(user=request.user).exists():
                    try:
                        org = Organization.objects.get(user=request.user)
                        UserOrg.objects.create(user=user, org=org)
                        UserRules.objects.create(user=user, full_access=form.cleaned_data.get('full_access'),
                                                 can_view_faculty=form.cleaned_data.get('can_view_faculty'),
                                                 can_add_faculty=form.cleaned_data.get('can_add_faculty'),
                                                 can_update_faculty=form.cleaned_data.get('can_update_faculty'),
                                                 can_delete_faculty=form.cleaned_data.get('can_delete_faculty'),
                                                 can_view_yonalish=form.cleaned_data.get('can_view_yonalish'),
                                                 can_add_yonalish=form.cleaned_data.get('can_add_yonalish'),
                                                 can_update_yonalish=form.cleaned_data.get('can_update_yonalish'),
                                                 can_delete_yonalish=form.cleaned_data.get('can_delete_yonalish'),
                                                 can_view_guruh=form.cleaned_data.get('can_view_guruh'),
                                                 can_add_guruh=form.cleaned_data.get('can_add_guruh'),
                                                 can_update_guruh=form.cleaned_data.get('can_update_guruh'),
                                                 can_delete_guruh=form.cleaned_data.get('can_delete_guruh'),
                                                 can_view_user=form.cleaned_data.get('can_view_user'),
                                                 can_add_user=form.cleaned_data.get('can_add_user'),
                                                 can_update_user=form.cleaned_data.get('can_update_user'),
                                                 can_delete_user=form.cleaned_data.get('can_delete_user'),
                                                 org=org,
                                                 username=username,
                                                 password=password,
                                                 )
                        if form.cleaned_data.get('full_access'):
                            user.user_permissions.add(Permission.objects.get(codename='add_userrulser'))
                            user.user_permissions.add(Permission.objects.get(codename='change_userrulser'))
                            user.user_permissions.add(Permission.objects.get(codename='delete_userrulser'))
                            user.user_permissions.add(Permission.objects.get(codename='view_userrulser'))
                            user.user_permissions.add(Permission.objects.get(codename='view_user'))
                            user.user_permissions.add(Permission.objects.get(codename='add_user'))
                            user.user_permissions.add(Permission.objects.get(codename='change_user'))
                            user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                            user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                            user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                            user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                            user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                            user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                            user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                            user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                            user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                            user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                            user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                            user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                            user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))

                        if form.cleaned_data.get('can_view_faculty'):
                            user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                        if form.cleaned_data.get('can_add_faculty'):
                            user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                        if form.cleaned_data.get('can_update_faculty'):
                            user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                        if form.cleaned_data.get('can_delete_faculty'):
                            user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                        if form.cleaned_data.get('can_view_yonalish'):
                            user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                        if form.cleaned_data.get('can_add_yonalish'):
                            user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                        if form.cleaned_data.get('can_update_yonalish'):
                            user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                        if form.cleaned_data.get('can_delete_yonalish'):
                            user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                        if form.cleaned_data.get('can_view_guruh'):
                            user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                        if form.cleaned_data.get('can_add_guruh'):
                            user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                        if form.cleaned_data.get('can_update_guruh'):
                            user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                        if form.cleaned_data.get('can_delete_guruh'):
                            user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                        if form.cleaned_data.get('can_view_user'):
                            user.user_permissions.add(Permission.objects.get(codename='view_user'))
                        if form.cleaned_data.get('can_add_user'):
                            user.user_permissions.add(Permission.objects.get(codename='add_user'))
                        if form.cleaned_data.get('can_update_user'):
                            user.user_permissions.add(Permission.objects.get(codename='change_user'))
                        if form.cleaned_data.get('can_delete_user'):
                            user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                        return redirect('userlist')
                    except:
                        user.delete()
                        messages.error(request, 'Form is not valid')
                        return redirect('userlist')

                if UserOrg.objects.filter(user=request.user).exists():
                    org = UserOrg.objects.get(user=request.user).org
                    UserOrg.objects.create(user=user, org=org)
                    UserRules.objects.create(user=user, full_access=form.cleaned_data.get('full_access'),
                                             can_view_faculty=form.cleaned_data.get('can_view_faculty'),
                                             can_add_faculty=form.cleaned_data.get('can_add_faculty'),
                                             can_update_faculty=form.cleaned_data.get('can_update_faculty'),
                                             can_delete_faculty=form.cleaned_data.get('can_delete_faculty'),
                                             can_view_yonalish=form.cleaned_data.get('can_view_yonalish'),
                                             can_add_yonalish=form.cleaned_data.get('can_add_yonalish'),
                                             can_update_yonalish=form.cleaned_data.get('can_update_yonalish'),
                                             can_delete_yonalish=form.cleaned_data.get('can_delete_yonalish'),
                                             can_view_guruh=form.cleaned_data.get('can_view_guruh'),
                                             can_add_guruh=form.cleaned_data.get('can_add_guruh'),
                                             can_update_guruh=form.cleaned_data.get('can_update_guruh'),
                                             can_delete_guruh=form.cleaned_data.get('can_delete_guruh'),
                                             can_view_user=form.cleaned_data.get('can_view_user'),
                                             can_add_user=form.cleaned_data.get('can_add_user'),
                                             can_update_user=form.cleaned_data.get('can_update_user'),
                                             can_delete_user=form.cleaned_data.get('can_delete_user'),
                                             org=org,
                                             username=username,
                                             password=password,
                                             )
                    if form.cleaned_data.get('full_access'):
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))

                    if form.cleaned_data.get('can_view_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                    if form.cleaned_data.get('can_add_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                    if form.cleaned_data.get('can_update_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                    if form.cleaned_data.get('can_delete_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                    if form.cleaned_data.get('can_view_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                    if form.cleaned_data.get('can_add_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                    if form.cleaned_data.get('can_update_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                    if form.cleaned_data.get('can_delete_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                    if form.cleaned_data.get('can_view_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                    if form.cleaned_data.get('can_add_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                    if form.cleaned_data.get('can_update_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                    if form.cleaned_data.get('can_delete_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                    if form.cleaned_data.get('can_view_user'):
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                    if form.cleaned_data.get('can_add_user'):
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                    if form.cleaned_data.get('can_update_user'):
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                    if form.cleaned_data.get('can_delete_user'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                    return redirect('userlist')
                return redirect('userlist')
            else:
                messages.error(request, 'Form is not valid')
                return redirect('userlist')


        else:
            form = UserRule()
            return render(request, 'baseapp/user_form.html', {'form': form})
    messages.error(request, 'You do not have permission to add user')
    return redirect('userlist')


@login_required(login_url='login')
def user_update(request, pk):
    if request.user.user_permissions.filter(codename='change_user').exists():
        global user
        user = User.objects.get(id=pk)

        if UserOrg.objects.filter(user=request.user).exists():

            user_rule = UserRules.objects.get(user=user)
            form = UserRuleUpdate(instance=user_rule)
            if request.method == 'POST':
                form = UserRule(request.POST, instance=user_rule)
                if form.is_valid():
                    form.save()
                    user.username = form.cleaned_data.get('username')
                    user.set_password(form.cleaned_data.get('password'))
                    user.user_permissions.clear()
                    if form.cleaned_data.get('full_access'):
                        user.user_permissions.add(Permission.objects.get(codename='add_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='change_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='view_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                    if form.cleaned_data.get('can_view_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                    if form.cleaned_data.get('can_add_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                    if form.cleaned_data.get('can_update_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                    if form.cleaned_data.get('can_delete_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                    if form.cleaned_data.get('can_view_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                    if form.cleaned_data.get('can_add_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                    if form.cleaned_data.get('can_update_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                    if form.cleaned_data.get('can_delete_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                    if form.cleaned_data.get('can_view_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                    if form.cleaned_data.get('can_add_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                    if form.cleaned_data.get('can_update_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                    if form.cleaned_data.get('can_delete_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                    if form.cleaned_data.get('can_view_user'):
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                    if form.cleaned_data.get('can_add_user'):
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                    if form.cleaned_data.get('can_update_user'):
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                    if form.cleaned_data.get('can_delete_user'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                    user.save()
                    return redirect('userlist')
            context = {'form': form}
            return render(request, 'baseapp/user_form.html', context)
        if Organization.objects.filter(user=request.user).exists():
            user_rule = UserRules.objects.get(user=user)
            form = UserRuleUpdate(instance=user_rule)
            if request.method == 'POST':
                form = UserRule(request.POST, instance=user_rule)
                if form.is_valid():
                    form.save()
                    user.username = form.cleaned_data.get('username')
                    user.set_password(form.cleaned_data.get('password'))
                    user.user_permissions.clear()
                    if form.cleaned_data.get('full_access'):
                        user.user_permissions.add(Permission.objects.get(codename='add_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='change_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='view_userrulser'))
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                    if form.cleaned_data.get('can_view_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='view_faculty'))
                    if form.cleaned_data.get('can_add_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='add_faculty'))
                    if form.cleaned_data.get('can_update_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='change_faculty'))
                    if form.cleaned_data.get('can_delete_faculty'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_faculty'))
                    if form.cleaned_data.get('can_view_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='view_yonalish'))
                    if form.cleaned_data.get('can_add_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='add_yonalish'))
                    if form.cleaned_data.get('can_update_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='change_yonalish'))
                    if form.cleaned_data.get('can_delete_yonalish'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_yonalish'))
                    if form.cleaned_data.get('can_view_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='view_guruh'))
                    if form.cleaned_data.get('can_add_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='add_guruh'))
                    if form.cleaned_data.get('can_update_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='change_guruh'))
                    if form.cleaned_data.get('can_delete_guruh'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_guruh'))
                    if form.cleaned_data.get('can_view_user'):
                        user.user_permissions.add(Permission.objects.get(codename='view_user'))
                    if form.cleaned_data.get('can_add_user'):
                        user.user_permissions.add(Permission.objects.get(codename='add_user'))
                    if form.cleaned_data.get('can_update_user'):
                        user.user_permissions.add(Permission.objects.get(codename='change_user'))
                    if form.cleaned_data.get('can_delete_user'):
                        user.user_permissions.add(Permission.objects.get(codename='delete_user'))
                    user.save()
                    return redirect('userlist')
                else:
                    messages.error(request, 'Form is not valid')
                    return redirect('userlist')
            context = {'form': form}
            return render(request, 'baseapp/user_form.html', context)
        else:
            messages.error(request, 'You do not have permission to update user')
            return redirect('userlist')
    messages.error(request, 'You do not have permission to update user')
    return redirect('userlist')


@login_required(login_url='login')
def user_delete(request, pk):
    if request.user.user_permissions.filter(codename='delete_user').exists():
        user = User.objects.get(id=pk)
        user.delete()
        return redirect('userlist')
    else:
        messages.error(request, 'You do not have permission to delete user')
        return redirect('userlist')


@login_required(login_url='login')
def emptypage(request):
    return render(request, 'pages/empty.html')


#
# @login_required(login_url='login')
# def generateexcel(request):
#     org = None
#     if UserOrg.objects.filter(user=request.user).exists():
#         org = UserOrg.objects.get(user=request.user).org
#     if Organization.objects.filter(user=request.user).exists():
#         org = Organization.objects.get(user=request.user)
#     if org:
#         if request.method == 'POST':
#             excel_files = generate_excel_files(org)
#
#             # Merge the Excel files
#             merged_output = merge_excel_files(excel_files)
#
#             # Set the response headers
#             response = HttpResponse(merged_output.getvalue(),
#                                     content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#             response['Content-Disposition'] = 'attachment; filename=merged_excel.xlsx'
#
#             return response
#         else:
#             return render(request, 'pages/generateexcel.html')
#     else:
#         messages.error(request, 'Your organization information is missing.')
#         return redirect('empty')


@background(schedule=60)
def asyncexcelgenarate(org):
    excel_files = generate_excel_files(org)
    merged_output = merge_excel_files(excel_files)
    response = HttpResponse(merged_output.getvalue(),
                            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=merged_excel.xlsx'
    return response


@login_required(login_url='login')
def generateexcel(request):
    org = None
    if UserOrg.objects.filter(user=request.user).exists():
        org = UserOrg.objects.get(user=request.user).org
    if Organization.objects.filter(user=request.user).exists():
        org = Organization.objects.get(user=request.user)
    if request.method == 'POST':
        if org:
            excel_files = generate_excel_files(org)
            merged_output = merge_excel_files(excel_files)
            response = HttpResponse(merged_output.getvalue(),
                                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename=merged_excel.xlsx'
            return response
        else:
            messages.error(request, 'Your organization information is missing.')
            return redirect('empty')
    else:
        return render(request, 'pages/generateexcel.html')
