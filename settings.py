import os
from django.utils.translation import gettext_lazy as _

# Idioma padrão
LANGUAGE_CODE = 'pt-br'

# Lista de idiomas permitidos
LANGUAGES = [
    ('pt-br', _('Português')),
    ('en', _('Inglês')),
]

# Localização onde os arquivos .po e .mo serão armazenados
# Geralmente criamos uma pasta 'locale' na raiz do projeto
LOCALE_PATHS = [
    os.path.join(BASE_DIR, 'locale'),
]

USE_I18N = True
USE_L10N = True
