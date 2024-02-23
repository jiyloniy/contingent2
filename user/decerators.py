from django.http import HttpResponse

from user.models import UserRulser


# write decorator for user rule
def user_rule_decorator(func):
    def wrapper(request, *args, **kwargs):
        user = request.user
        user_rule = UserRulser.objects.get(user=user)
        if user_rule.can_view:
            return func(request, *args, **kwargs)
        else:
            return HttpResponse('You have no permission to view this page')

    return wrapper
