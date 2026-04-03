# Snel aan de slag met de Urenstaat-Tool 🚀

Hoi! Met deze tool bespaar je jezelf het gedoe van urenstaten handmatig omzetten. In een paar tellen heb je een mooie PDF en Excel klaarstaan in Outlook.

## Hoe werkt het?

1. **Eenmalige Setup:**
   - Dubbelklik op `setup_mac.command` (macOS) of `setup_windows.bat` (Windows).
   - Vul je naam, klantnaam en de map in waar je de bestanden wilt bewaren (bijv. je OneDrive map).
   - Er verschijnt een icoontje **"Process Timesheet"** op je bureaublad.

2. **Uren verwerken:**
   - Download je uren-export (.csv) van Float/People.
   - Sleep dit bestand simpelweg op het **"Process Timesheet"** icoontje op je bureaublad.
   - De tool doet de rest: hij maakt de PDF/Excel, zet ze in de juiste map en opent een concept-mail in Outlook.

## Goed om te weten
- Je originele CSV wordt na verwerking automatisch verplaatst naar de map `converted/` (zo blijft je bureaublad netjes).
- De PDF en Excel komen in de map te staan die je tijdens de setup hebt gekozen.
- Wil je iets aanpassen (zoals je klantnaam)? Run de setup gewoon opnieuw of pas `config.json` aan.

Succes! ✨
