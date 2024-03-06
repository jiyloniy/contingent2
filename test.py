import os
import re

import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

django.setup()
from user.models import Faculty, Budjet, Shartnoma, Organization, Yonalish, Guruh

from exportexcel import exporttoexcel
org = Organization.objects.filter(name='kiuf').first()
exporttoexcel(org)