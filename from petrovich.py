from petrovich.main import Petrovich
from petrovich.enums import Case, Gender

petrovich = Petrovich()

surname = "Иванов"
genitive = petrovich.lastname(surname, case=Case.GENITIVE, gender=Gender.MALE)
dative   = petrovich.lastname(surname, case=Case.DATIVE, gender=Gender.MALE)

print("Родительный падеж:", genitive)  # Ожидается: Иванова
print("Дательный падеж:", dative)      # Ожидается: Иванову
