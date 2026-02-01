# 🌊 Wildwind Rumstillgänglighet

Automatisk uppdatering av rumstillgänglighet för Wildwind-resor.

Sidan visar vilka rum som är lediga varje vecka (lördag-lördag) baserat på Stephs boknings-Excel i Dropbox.

## 🚀 Så här sätter du upp det

### Steg 1: Skapa ett GitHub-konto (om du inte har)
1. Gå till [github.com](https://github.com)
2. Klicka "Sign up" och skapa ett konto

### Steg 2: Skapa ett nytt repository
1. Klicka på **+** uppe till höger → **New repository**
2. Fyll i:
   - **Repository name:** `wildwind-availability`
   - **Description:** Wildwind rumstillgänglighet 2026
   - ✅ Kryssa i **Public**
   - ✅ Kryssa i **Add a README file**
3. Klicka **Create repository**

### Steg 3: Ladda upp filerna
1. I ditt nya repository, klicka **Add file** → **Upload files**
2. Dra och släpp dessa filer:
   - `update_availability.py`
3. Skriv "Initial setup" som commit message
4. Klicka **Commit changes**

### Steg 4: Skapa workflow-mappen
1. Klicka **Add file** → **Create new file**
2. I filnamnet, skriv: `.github/workflows/update.yml`
3. Klistra in innehållet från `update.yml`-filen
4. Klicka **Commit changes**

### Steg 5: Aktivera GitHub Pages
1. Gå till **Settings** (kugghjulet)
2. Scrolla ner till **Pages** i vänstermenyn
3. Under **Source**, välj:
   - Branch: `main`
   - Folder: `/ (root)`
4. Klicka **Save**
5. Vänta 1-2 minuter, sedan visas din URL: `https://DITTANVÄNDARNAMN.github.io/wildwind-availability/`

### Steg 6: Ge Actions rätt att pusha
1. Gå till **Settings** → **Actions** → **General**
2. Scrolla ner till **Workflow permissions**
3. Välj **Read and write permissions**
4. Klicka **Save**

### Steg 7: Kör första uppdateringen manuellt
1. Gå till **Actions**-fliken
2. Klicka på **Update Wildwind Availability**
3. Klicka **Run workflow** → **Run workflow**
4. Vänta tills den blir grön ✅
5. Din sida är nu live!

---

## 🔗 Länka från travel.seafari.se

När allt fungerar kan du antingen:

**Alternativ A: Redirect**
Lägg till en redirect i din webbserver/hosting:
```
travel.seafari.se/i/wildwind-bokningsforfragan → https://DITT.github.io/wildwind-availability/
```

**Alternativ B: iFrame**
Bädda in på din sida:
```html
<iframe src="https://DITT.github.io/wildwind-availability/" 
        style="width:100%; height:100vh; border:none;">
</iframe>
```

**Alternativ C: Egen domän på GitHub Pages**
1. I repository Settings → Pages
2. Under "Custom domain", skriv: `wildwind.seafari.se` (eller liknande)
3. Lägg till en CNAME-post i din DNS som pekar till `DITT.github.io`

---

## ⏰ Uppdateringsschema

Sidan uppdateras automatiskt varje dag kl 06:00 UTC (07:00/08:00 svensk tid).

Du kan också trigga en uppdatering manuellt:
1. Gå till Actions
2. Klicka på workflowen
3. Klicka "Run workflow"

---

## 📝 Felsökning

**Scriptet hittar inte filen?**
- Kontrollera att Dropbox-länken fortfarande fungerar
- Länken måste vara delad så "alla med länken" kan se den

**Actions misslyckas?**
- Kolla i Actions-loggen för felmeddelanden
- Kontrollera att Workflow permissions är "Read and write"

**Sidan uppdateras inte?**
- GitHub Pages kan ta några minuter att uppdatera
- Prova att rensa webbläsarens cache (Ctrl+Shift+R)

---

## 💡 Anpassa

Vill du ändra något? Redigera `update_availability.py`:

- **Andra rum?** Ändra `ALLOWED_ROWS`-listan
- **Annan Dropbox-länk?** Ändra `DROPBOX_URL`
- **Annat utseende?** Ändra CSS i `generate_html()`-funktionen

---

Skapad med ❤️ av Claude för Seafari
