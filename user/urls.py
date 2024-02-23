from django.urls import path

from user.views import login_view, logout_view, faculty_list, add_faculty, faculty_update, faculty_delete, \
    yonalish_list, yonalish_create, yonalish_update, yonalish_delete, guruh_list, guruh_create, guruh_update, \
    guruh_delete

urlpatterns = [
    path('login/', login_view, name='login'),
    path('logout/', logout_view, name='logout'),
    path('', faculty_list, name='dashboard'),
    path('add-faculty/', add_faculty, name='facultycreate'),
    path('facultyupdate/<int:pk>/', faculty_update, name='facultyupdate'),
    path('facultydelete/<int:pk>/', faculty_delete, name='facultydelete'),
    path('yonalish/', yonalish_list, name='yonalishlar'),
    path('yonalishcreate/', yonalish_create, name='yonalishcreate'),
    path('yonalishupdate/<int:pk>/', yonalish_update, name='yonalishupdate'),
    path('yonalishdelete/<int:pk>/', yonalish_delete, name='yonalishdelete'),
    path('guruh/', guruh_list, name='guruh'),
    path('guruhcreate/', guruh_create, name='group_create'),
    path('guruhupdate/<int:pk>/', guruh_update, name='guruhupdate'),
    path('guruhdelete/<int:pk>/', guruh_delete, name='guruhdelete'),

]
