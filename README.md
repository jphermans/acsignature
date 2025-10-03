# ✨ E-mail Signature Generator

<p align="center">
  <img src="https://img.shields.io/badge/version-1.0.0-0092BC?style=for-the-badge" />
  <img src="https://img.shields.io/badge/status-stable-2F363A?style=for-the-badge" />
</p>

---

## 📌 Over dit project
Een eenvoudige tool om automatisch professionele e-mailhandtekeningen te genereren.  
Gebouwd met **HTML, CSS en JavaScript**.  

De generator:
- ✅ Maakt automatisch **.htm, .rtf en .txt** versies aan (compatibel met Outlook).  
- ✅ Zorgt dat alle verplichte velden (Naam, Functie, E-mail, Mobiel) ingevuld zijn.  
- ✅ Controleert dat het e-mailadres eindigt op `@atlascopco.com`.  
- ✅ Exporteert direct naar de Outlook **Signatures** map (handmatig kopiëren).  
- ✅ Geeft duidelijke foutmeldingen bij ontbrekende velden.  

---

## 🎨 Layout & Kleuren
De tool gebruikt een strak en clean design met herkenbare kleuren:

- **Hoofdaccentkleur:** ![#0092BC](https://via.placeholder.com/15/0092BC/000000?text=+) `#0092BC`  
- **Tekstkleur donkergrijs:** ![#2F363A](https://via.placeholder.com/15/2F363A/000000?text=+) `#2F363A`  
- **Lijntjes lichtgrijs:** ![#E1D6CE](https://via.placeholder.com/15/E1D6CE/000000?text=+) `#E1D6CE`  

---

## ⚙️ Installatie & Gebruik

### 1. Open de generator
Download of clone dit project en open het bestand **index.html** in je browser.

### 2. Vul je gegevens in
- Naam  
- Functie  
- E-mail (moet eindigen op `@atlascopco.com`)  
- Mobiel (verplicht)  
- Telefoon (optioneel)  

### 3. Genereer je handtekening
Klik op **Genereer Handtekening**.  
Je ziet meteen een preview.

### 4. Exporteer naar Outlook
Klik op **Exporteer naar Outlook** → er worden 3 bestanden aangemaakt:
- `AtlasCopco.htm`  
- `AtlasCopco.rtf`  
- `AtlasCopco.txt`  

### 5. Zet de bestanden in Outlook
- Open de map:  
  ```bash
  %appdata%\Microsoft\Signatures
