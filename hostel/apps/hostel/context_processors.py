from apps.hostel.models import Notification, RoleModulePermission


def notifications_ctx(request):
    """Inject unread notification count into every template."""
    if request.user.is_authenticated:
        unread_count = Notification.objects.filter(
            recipient=request.user, is_read=False
        ).count()
        return {'unread_notifications_count': unread_count}
    return {'unread_notifications_count': 0}


def role_permissions(request):
    """Inject visible_modules set so the sidebar can hide/show links per role."""
    if request.user.is_authenticated:
        return {'visible_modules': RoleModulePermission.get_for_role(request.user.role)}
    return {'visible_modules': set()}
