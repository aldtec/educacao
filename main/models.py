from django.db import models

# Create your models here.
class Docente(models.Model):
    nome      = models.CharField(max_length=50)
    rf_vinc   = models.CharField(max_length=12)
    qpe       = models.CharField(max_length=4, blank=True)
    cargo     = models.CharField(max_length=50, blank=True, default='Professor Educ. Infantil e Ens. Fund. I')
    regencia  = models.CharField(max_length=20, blank=True)
    hor_col   = models.CharField(max_length=30, blank=True)
    turma     = models.CharField(max_length=6, blank=True)
    horario   = models.CharField(max_length=80, blank=True)
    JORNADAS  = (('J', 'JEIF'),('D', 'JBD'),('B', 'JB'),)
    jornada   = models.CharField(max_length=1, choices=JORNADAS, default='D')

    def __str__(self):
    	return self.nome

class Calendario(models.Model):
    descricao = models.CharField(max_length=30)
    data      = models.DateField()
    JORNADAS  = (
    	('F', 'FERIADO'),
    	('R', 'RECESSO'),
    	('P', 'PONTO FACULTATIVO'),
    	('L', 'DIA NÃO LETIVO'),
    	)
    observ   = models.CharField(max_length=1, choices=JORNADAS, default='F')

    def __str__(self):
    	return self.descricao + " - " + self.data.strftime("%d/%m/%Y")