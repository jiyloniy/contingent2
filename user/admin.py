from django.contrib import admin
from django.contrib.auth.models import Permission
from guardian.admin import GuardedModelAdmin
# Register your models here.
from .models import UserOrg, Organization,Faculty, Yonalish, Shartnoma, Budjet, Guruh

admin.site.register(Organization, GuardedModelAdmin)
# permission = Permission.objects.get(codename='can_view')

admin.site.register(Permission)
class UserProfileAdmin(admin.ModelAdmin):
    list_display = ('organization',)
    search_fields = ('organization',)


class UserOrgAdmin(admin.ModelAdmin):
    list_display = ('org', 'user')
    search_fields = ('org', 'user')
    list_filter = ('org', 'user')


admin.site.register(UserOrg, UserOrgAdmin)
admin.site.register(Faculty)
admin.site.register(Yonalish)
admin.site.register(Shartnoma)
admin.site.register(Budjet)


