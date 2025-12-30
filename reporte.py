import time
import requests
from msal import ConfidentialClientApplication
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
import tempfile
import os
from PIL import Image

# === DATOS ===
tenant_id = "00000000-0000-0000-0000-000000000000"
client_id = "00000000-0000-0000-0000-000000000000"
client_secret = "asdfghjklñ"
workspace_id = "00000000-0000-0000-0000-000000000000"
report_id = "00000000-0000-0000-0000-000000000000"

# === 1. TOKEN DEL SERVICE PRINCIPAL ===
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://analysis.windows.net/powerbi/api/.default"]

app = ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret
)

token_response = app.acquire_token_for_client(scopes=scope)
access_token = token_response["access_token"]
print("Token de servicio obtenido")

# === 2. TOKEN DE EMBED ===
generate_token_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{report_id}/GenerateToken"
body = {"accessLevel": "View"}

embed_token_response = requests.post(
    generate_token_url,
    headers={
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    },
    json=body
)

if embed_token_response.status_code != 200:
    print("Error al generar token de embed:")
    print(embed_token_response.text)
    exit()

embed_token = embed_token_response.json()["token"]
print("Token de embed generado correctamente")

# === 3. HTML TEMPORAL ===
html_content = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<script src="https://cdn.jsdelivr.net/npm/powerbi-client@2.19.0/dist/powerbi.js"></script>
</head>
<body style="margin:0;overflow:hidden;">
<div id="reportContainer" style="width:100vw; height:100vh;"></div>
<script>
var models = window['powerbi-client'].models;

// Configuracion de embed minimalista y limpia
var embedConfiguration = {
    type: 'report',
    tokenType: models.TokenType.Embed,
    accessToken: '__EMBED_TOKEN__',
    embedUrl: 'https://app.powerbi.com/reportEmbed?reportId=__REPORT_ID__&groupId=__WORKSPACE_ID__',
    settings: {
        panes: {
            filters: { visible: false },
            pageNavigation: { visible: false }
        },
        layoutType: models.LayoutType.Custom,
        customLayout: { displayOption: models.DisplayOption.FitToPage }
    }
};

var reportContainer = document.getElementById('reportContainer');
var report = powerbi.embed(reportContainer, embedConfiguration);

(async function () {
    try {
        // Esperar a que el reporte cargue
        await new Promise(resolve => report.on('loaded', resolve));

        const pages = await report.getPages();

        for (let i = 0; i < pages.length; i++) {
            const page = pages[i];
            await page.setActive();                 // Activar la pagina
            await new Promise(r => setTimeout(r, 1000)); // Espera minima a que cargue contenido
            document.title = "PAGE_RENDERED_" + i;  // Avisar a Selenium
            await new Promise(r => setTimeout(r, 1000));
        }

        document.title = "ALL_PAGES_RENDERED"; // Todas las paginas listas
    } catch (err) {
        console.error("Error durante renderizado:", err);
        document.title = "PAGE_ERROR";
    }
})();
</script>
</body>
</html>
"""

html_content = (
    html_content
    .replace("__EMBED_TOKEN__", embed_token)
    .replace("__REPORT_ID__", report_id)
    .replace("__WORKSPACE_ID__", workspace_id)
)

# === 4. GUARDAR HTML TEMPORAL ===
temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html")
temp_file.write(html_content.encode("utf-8"))
temp_file.close()
html_path = f"file:///{temp_file.name.replace(os.sep, '/')}"

# === 5. SELENIUM CONFIGURACIoN ===
chrome_options = Options()
chrome_options.add_argument("--window-size=1936,1120")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--force-device-scale-factor=1")
chrome_options.add_argument("--high-dpi-support=1")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
print("Cargando reporte embebido...")
driver.get(html_path)
wait = WebDriverWait(driver, 600)

# === 6. CAPTURAR CADA PaGINA COMO IMAGEN ===
capturas = []
page_index = 0
last_title = ""

print("Esperando a que las paginas se rendericen...")

while True:
    wait.until(lambda d: any(x in d.title for x in ["PAGE_RENDERED_", "ALL_PAGES_RENDERED", "PAGE_ERROR"]))
    title = driver.title

    if title.startswith("PAGE_RENDERED_") and title != last_title:
        filename = f"powerbi_page_{page_index}.png"
        driver.save_screenshot(filename)
        capturas.append(filename)
        print(f"Captura guardada: {filename}")
        page_index += 1
        last_title = title

    elif "ALL_PAGES_RENDERED" in title:
        print("Todas las paginas renderizadas correctamente.")
        break

    elif "PAGE_ERROR" in title:
        print("Error durante renderizado.")
        break

driver.quit()
os.unlink(temp_file.name)

# === 7. UNIR IMaGENES EN PDF ===
if not capturas:
    print("No se generaron capturas. Revisa el renderizado del informe.")
    exit()

images = [Image.open(img).convert("RGB") for img in capturas]

# Cada imagen sera una pagina del PDF
images[0].save(
    "powerbi_z1.pdf",
    save_all=True,
    append_images=images[1:],
)
print("Dashboard completo guardado como powerbi_z1.pdf")

# === 8. LIMPIAR ARCHIVOS TEMPORALES ===
for img in capturas:
    os.remove(img)
