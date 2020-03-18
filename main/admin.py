from django.contrib import admin

from .models import Docente, Calendario

# Register your models here.
admin.site.register(Docente)
admin.site.register(Calendario)