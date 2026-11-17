from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Cria usuários de demonstração para admin e viewer."

    def handle(self, *args, **options):
        User = get_user_model()

        admin_user, admin_created = User.objects.get_or_create(
            username="admin",
            defaults={"email": "admin@jangle.local", "role": User.ROLE_ADMIN, "is_staff": True, "is_superuser": True},
        )
        if admin_created:
            admin_user.set_password("admin123")
            admin_user.save()
        self.stdout.write(self.style.SUCCESS("Usuário admin criado/confirmado: admin / admin123"))

        viewer_user, viewer_created = User.objects.get_or_create(
            username="viewer",
            defaults={"email": "viewer@jangle.local", "role": User.ROLE_VIEWER, "is_staff": False, "is_superuser": False},
        )
        if viewer_created:
            viewer_user.set_password("viewer123")
            viewer_user.save()
        self.stdout.write(self.style.SUCCESS("Usuário viewer criado/confirmado: viewer / viewer123"))
