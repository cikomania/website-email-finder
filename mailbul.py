import re
import time
import requests
import pandas as pd

from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

from concurrent.futures import ThreadPoolExecutor, as_completed
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# OCR
try:
    import pytesseract
    from PIL import Image
    from io import BytesIO
    OCR_VAR = True
except ImportError:
    OCR_VAR = False


##############################################################################
# AYARLAR
##############################################################################

INPUT = "firmalar.xlsx"
OUTPUT = "firmalar_mailli.xlsx"

# Bir firmanın tamamını taramak için maksimum süre
MAX_FIRMA_SURE = 18

# Ana sayfa dışında maksimum önemli sayfa
MAX_CONTACT_PAGE = 5

# Aynı anda kaç firma çalışsın?
MAX_WORKERS = 20

# OCR açık/kapalı
OCR_AKTIF = True

# Bir sayfada OCR yapılabilecek maksimum resim
MAX_OCR_IMAGE = 4

# OCR için minimum resim genişlik/yükseklik
MIN_IMAGE_WIDTH = 150
MIN_IMAGE_HEIGHT = 50


##############################################################################
# EMAIL REGEX
##############################################################################

EMAIL_REGEX = (
    r"\b[A-Za-z0-9._%+\-]+"
    r"@"
    r"[A-Za-z0-9.\-]+"
    r"\.[A-Za-z]{2,24}\b"
)

EMAIL_PATTERN = re.compile(
    EMAIL_REGEX,
    re.IGNORECASE
)


##############################################################################
# HEADERS
##############################################################################

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/138.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
    "Connection": "keep-alive"
}


##############################################################################
# MAIL ÖNCELİĞİ
##############################################################################

ONCELIK = [
    "info@",
    "iletisim@",
    "contact@",
    "mail@",
    "kurumsal@",
    "satis@",
    "sales@",
    "destek@",
    "office@",
    "admin@"
]


##############################################################################
# SAHTE / İSTENMEYEN MAIL KELİMELERİ
##############################################################################

RED = [
    "noreply",
    "no-reply",
    "donotreply",
    "do-not-reply",

    "career",
    "kariyer",
    "hr@",
    "ik@",

    "wordpress",
    "mysite",

    "sentry",
    "@sentry.io",

    "cloudflare",
    "cfemail",

    "facebook",
    "instagram",
    "linkedin",
    "twitter",
    "youtube",

    "%20",

    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".svg",
    ".webp"
]


##############################################################################
# SAHTE KULLANICI ADLARI
##############################################################################

SAHTE_KULLANICILAR = {
    "test",
    "demo",
    "example",
    "examples",
    "username",
    "yourname",
    "your-name",
    "name",
    "firstname",
    "lastname",
    "user",
    "domain",
    "sitem",
    "site",
    "website",
    "web",
    "email",
    "mail",
    "abc",
    "xyz",
    "xxx",
    "sample",
    "deneme",
    "ornek",
    "örnek"
}


##############################################################################
# SAHTE DOMAINLER
##############################################################################

SAHTE_DOMAINLER = {
    "example.com",
    "example.org",
    "example.net",

    "domain.com",
    "domain.org",
    "domain.net",

    "test.com",
    "test.org",
    "test.net",

    "localhost",
    "localhost.localdomain"
}


##############################################################################
# İLETİŞİM SAYFASI ANAHTARLARI
##############################################################################

CONTACT_KEYWORDS = [
    "iletisim",
    "iletişim",
    "contact",
    "contact-us",
    "contactus",
    "contact_us",

    "hakkimizda",
    "hakkımızda",
    "about",

    "kurumsal",
    "corporate",
    "company",
    "firma",

    "ulasin",
    "ulaşın",
    "bize-ulasin",
    "bize-ulaşın",
    "bizeulasin",

    "reach",
    "communication",

    "office",
    "adres",
    "address"
]


##############################################################################
# OCR İÇİN ANAHTARLAR
##############################################################################

OCR_KEYWORDS = [
    "iletisim",
    "iletisim",
    "contact",
    "mail",
    "email",
    "e-mail",
    "adres",
    "address",
    "footer",
    "communication"
]


##############################################################################
# SESSION
##############################################################################

retry = Retry(
    total=0,
    connect=0,
    read=0,
    redirect=3,
    status=0
)

adapter = HTTPAdapter(
    max_retries=retry,
    pool_connections=30,
    pool_maxsize=30
)

session = requests.Session()

session.mount(
    "http://",
    adapter
)

session.mount(
    "https://",
    adapter
)

session.headers.update(HEADERS)


##############################################################################
# EXCEL TEXT
##############################################################################

def excel_text(text):

    if not isinstance(text, str):
        return text

    if text.startswith("+"):
        return "'" + text

    return text


##############################################################################
# URL TEMİZLE
##############################################################################

def url_temizle(site):

    if pd.isna(site):
        return ""

    site = str(site).strip()

    if not site:
        return ""

    if site.lower() == "nan":
        return ""

    # Markdown link
    markdown_match = re.search(
        r'\]\((https?://[^)]+)\)',
        site,
        re.IGNORECASE
    )

    if markdown_match:
        site = markdown_match.group(1)

    site = site.replace("[", "")
    site = site.replace("]", "")

    site = site.strip()

    site = re.sub(
        r"^https?://",
        "",
        site,
        flags=re.IGNORECASE
    )

    site = site.strip().strip("/")

    return site


##############################################################################
# OLASI URL'LER
##############################################################################

def olasi_urller(site):

    site = url_temizle(site)

    if not site:
        return []

    site = site.split("/")[0]

    if site.lower().startswith("www."):
        kok = site[4:]
    else:
        kok = site

    if not kok:
        return []

    if any(
        x in kok
        for x in ["[", "]", "(", ")", " "]
    ):
        return []

    urller = [
        f"https://{kok}",
        f"https://www.{kok}",
        f"http://{kok}",
        f"http://www.{kok}"
    ]

    return list(
        dict.fromkeys(urller)
    )


##############################################################################
# URL GEÇERLİ Mİ?
##############################################################################

def gecerli_url(url):

    try:

        parsed = urlparse(url)

        if parsed.scheme not in (
            "http",
            "https"
        ):
            return False

        if not parsed.netloc:
            return False

        if any(
            x in parsed.netloc
            for x in ["[", "]", "(", ")", " "]
        ):
            return False

        return True

    except Exception:
        return False


##############################################################################
# DOMAIN ÇIKAR
##############################################################################

def domain_cikar(url):

    try:

        hostname = urlparse(url).hostname

        if not hostname:
            return ""

        hostname = hostname.lower()

        if hostname.startswith("www."):
            hostname = hostname[4:]

        return hostname

    except Exception:
        return ""


##############################################################################
# MAIL NORMALİZE
##############################################################################

def mail_normalize(mail):

    if not isinstance(mail, str):
        return ""

    mail = mail.strip().lower()

    mail = re.sub(
        r"^mailto:",
        "",
        mail,
        flags=re.IGNORECASE
    )

    mail = mail.split("?")[0]

    mail = mail.strip(
        ".,;:()[]{}<>\\\"' "
    )

    return mail


##############################################################################
# OCR / GİZLİ MAIL FORMATLARINI NORMALİZE ET
##############################################################################

def ocr_mail_normalize(text):

    if not isinstance(text, str):
        return ""

    text = text.replace(
        "[at]",
        "@"
    )

    text = text.replace(
        "(at)",
        "@"
    )

    text = text.replace(
        "{at}",
        "@"
    )

    text = re.sub(
        r"\s+at\s+",
        "@",
        text,
        flags=re.IGNORECASE
    )

    text = text.replace(
        "[dot]",
        "."
    )

    text = text.replace(
        "(dot)",
        "."
    )

    text = text.replace(
        "{dot}",
        "."
    )

    return text


##############################################################################
# DOMAIN BENZERLİĞİ
##############################################################################

def domain_benzer_mi(mail_domain, site_domain):

    if not mail_domain or not site_domain:
        return False

    mail_domain = mail_domain.lower()
    site_domain = site_domain.lower()

    if mail_domain == site_domain:
        return True

    # www farkı
    if mail_domain.startswith("www."):
        mail_domain = mail_domain[4:]

    if site_domain.startswith("www."):
        site_domain = site_domain[4:]

    if mail_domain == site_domain:
        return True

    # Alt domain
    if mail_domain.endswith(
        "." + site_domain
    ):
        return True

    return False


##############################################################################
# MAIL TEMİZLE
##############################################################################

def temizle(mailler):

    sonuc = []

    for mail in mailler:

        if not isinstance(mail, str):
            continue

        mail = ocr_mail_normalize(mail)
        mail = mail_normalize(mail)

        if not mail:
            continue

        if mail.count("@") != 1:
            continue

        try:

            kullanici, domain = mail.split("@")

        except Exception:
            continue

        if not kullanici or not domain:
            continue

        if "." not in domain:
            continue

        if len(kullanici) < 2:
            continue

        if len(kullanici) > 64:
            continue

        if len(domain) > 253:
            continue

        if any(
            x in mail
            for x in RED
        ):
            continue

        if kullanici in SAHTE_KULLANICILAR:
            continue

        if domain in SAHTE_DOMAINLER:
            continue

        # Domain'in de açıkça örnek olduğunu yakala
        domain_parcalari = domain.split(".")

        if any(
            parca in {
                "example",
                "exampledomain",
                "my-domain",
                "yourdomain",
                "domain",
                "test",
                "localhost",
                "sitem"
            }
            for parca in domain_parcalari
        ):
            continue

        # Hash benzeri kullanıcı
        if len(kullanici) > 25:

            rakam = sum(
                c.isdigit()
                for c in kullanici
            )

            if rakam > 5:
                continue

        # Uzun hex hash
        if re.fullmatch(
            r"[a-f0-9]{20,}",
            kullanici
        ):
            continue

        # Kullanıcı tamamen rakamsa
        if re.fullmatch(
            r"\d+",
            kullanici
        ):
            continue

        # Çok anlamsız tekrarlar
        if re.fullmatch(
            r"(.)\1{5,}",
            kullanici
        ):
            continue

        if mail not in sonuc:
            sonuc.append(mail)

    return sonuc


##############################################################################
# MAIL PUANLA
##############################################################################

def mail_puanla(mail, site_domain, kaynak="html"):

    puan = 0

    try:

        kullanici, domain = mail.lower().split("@")

    except Exception:
        return -999

    ##########################################################################
    # ŞİRKET DOMAINİ
    ##########################################################################

    if domain_benzer_mi(
        domain,
        site_domain
    ):

        puan += 100

    else:

        # Gmail / Hotmail gibi kişisel domainler
        if domain in {
            "gmail.com",
            "hotmail.com",
            "hotmail.com.tr",
            "outlook.com",
            "yahoo.com",
            "yahoo.com.tr",
            "icloud.com",
            "live.com"
        }:

            puan += 10

        else:

            puan += 25

    ##########################################################################
    # KURUMSAL MAIL ÖNCELİĞİ
    ##########################################################################

    for sira, tercih in enumerate(
        ONCELIK
    ):

        if mail.startswith(tercih):

            puan += 60 - (
                sira * 4
            )

            break

    ##########################################################################
    # GENEL MAILLER
    ##########################################################################

    if kullanici in {
        "info",
        "iletisim",
        "contact",
        "mail",
        "kurumsal",
        "satis",
        "sales",
        "office",
        "admin",
        "destek"
    }:

        puan += 25

    ##########################################################################
    # KAYNAK
    ##########################################################################

    if kaynak == "mailto":
        puan += 35

    elif kaynak == "contact":
        puan += 30

    elif kaynak == "footer":
        puan += 25

    elif kaynak == "attribute":
        puan += 20

    elif kaynak == "javascript":
        puan += 15

    elif kaynak == "ocr":
        puan += 15

    elif kaynak == "html":
        puan += 10

    ##########################################################################
    # RED KELİMELERİ
    ##########################################################################

    if any(
        x in mail
        for x in [
            "noreply",
            "no-reply",
            "donotreply",
            "career",
            "kariyer",
            "wordpress",
            "sentry",
            "cloudflare"
        ]
    ):

        puan -= 100

    ##########################################################################
    # TEST / ÖRNEK
    ##########################################################################

    if (
        kullanici in SAHTE_KULLANICILAR
    ):

        puan -= 150

    return puan


##############################################################################
# EN İYİ MAİLİ SEÇ
##############################################################################

def en_iyi_mail_sec(
    mailler,
    site_domain,
    mail_kaynaklari=None
):

    mailler = temizle(
        mailler
    )

    if not mailler:
        return ""

    if mail_kaynaklari is None:
        mail_kaynaklari = {}

    skorlar = []

    for mail in mailler:

        kaynak = mail_kaynaklari.get(
            mail,
            "html"
        )

        skor = mail_puanla(
            mail,
            site_domain,
            kaynak
        )

        skorlar.append(
            (
                skor,
                mail
            )
        )

    skorlar.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return skorlar[0][1]


##############################################################################
# HTML İÇİNDEN MAIL ÇIKAR
##############################################################################

def html_mail_bul(
    text,
    soup=None,
    kaynak="html"
):

    bulunan = []

    if not text:
        return bulunan

    ##########################################################################
    # Normal HTML
    ##########################################################################

    bulunan.extend(
        EMAIL_PATTERN.findall(
            text
        )
    )

    if soup is None:
        return bulunan

    ##########################################################################
    # MAILTO
    ##########################################################################

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        )

        if href.lower().startswith(
            "mailto:"
        ):

            mail = (
                href
                .replace(
                    "mailto:",
                    "",
                    1
                )
                .split("?")[0]
            )

            bulunan.append(
                mail
            )

    ##########################################################################
    # TÜM ATTRIBUTE'LAR
    ##########################################################################

    for tag in soup.find_all(True):

        for attr, value in tag.attrs.items():

            if isinstance(
                value,
                list
            ):

                value = " ".join(
                    str(x)
                    for x in value
                )

            if not isinstance(
                value,
                str
            ):

                continue

            bulunan.extend(
                EMAIL_PATTERN.findall(
                    value
                )
            )

    ##########################################################################
    # OCR FORMATLARI
    ##########################################################################

    normalize_text = ocr_mail_normalize(
        text
    )

    bulunan.extend(
        EMAIL_PATTERN.findall(
            normalize_text
        )
    )

    return bulunan


##############################################################################
# SAYFADAN MAIL TOPLA
##############################################################################

def sayfadan_mail_topla(
    response,
    mail_kaynaklari,
    kaynak="html"
):

    bulunan = set()

    try:

        text = response.text

        soup = BeautifulSoup(
            text,
            "html.parser"
        )

        mailler = html_mail_bul(
            text,
            soup
        )

        for mail in mailler:

            temiz = mail_normalize(
                mail
            )

            if temiz:

                bulunan.add(
                    temiz
                )

                if kaynak == "contact":
                    mail_kaynaklari[
                        temiz
                    ] = "contact"

                else:
                    mail_kaynaklari.setdefault(
                        temiz,
                        kaynak
                    )

        ######################################################################
        # FOOTER
        ######################################################################

        footer = soup.find(
            "footer"
        )

        if footer:

            footer_text = str(
                footer
            )

            footer_mailler = (
                EMAIL_PATTERN.findall(
                    footer_text
                )
            )

            for mail in footer_mailler:

                temiz = mail_normalize(
                    mail
                )

                if temiz:

                    bulunan.add(
                        temiz
                    )

                    mail_kaynaklari[
                        temiz
                    ] = "footer"

        ######################################################################
        # JAVASCRIPT
        ######################################################################

        for script in soup.find_all(
            "script"
        ):

            script_text = script.string

            if not script_text:
                script_text = script.get_text(
                    " ",
                    strip=False
                )

            if not script_text:
                continue

            script_text = ocr_mail_normalize(
                script_text
            )

            js_mailler = (
                EMAIL_PATTERN.findall(
                    script_text
                )
            )

            for mail in js_mailler:

                temiz = mail_normalize(
                    mail
                )

                if temiz:

                    bulunan.add(
                        temiz
                    )

                    mail_kaynaklari[
                        temiz
                    ] = "javascript"

    except Exception:
        pass

    return bulunan


##############################################################################
# İLETİŞİM LİNKLERİNİ BUL
##############################################################################

def iletisim_linklerini_bul(
    soup,
    base_url
):

    adaylar = []

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:
            continue

        if href.startswith(
            "#"
        ):
            continue

        if href.lower().startswith(
            "mailto:"
        ):
            continue

        link_text = a.get_text(
            " ",
            strip=True
        )

        birlikte = (
            href.lower()
            + " "
            + link_text.lower()
        )

        skor = 0

        for keyword in CONTACT_KEYWORDS:

            if keyword in birlikte:

                skor += 10

        if not skor:
            continue

        link = urljoin(
            base_url,
            href
        )

        if not gecerli_url(
            link
        ):
            continue

        # Aynı domain dışına çıkma
        base_domain = domain_cikar(
            base_url
        )

        link_domain = domain_cikar(
            link
        )

        if (
            base_domain
            and link_domain
            and base_domain != link_domain
        ):
            continue

        adaylar.append(
            (
                skor,
                link
            )
        )

    # En güçlü linkler önce
    adaylar.sort(
        key=lambda x: x[0],
        reverse=True
    )

    sonuc = []

    for _, link in adaylar:

        if link not in sonuc:
            sonuc.append(
                link
            )

    return sonuc


##############################################################################
# OCR İÇİN UYGUN RESİMLERİ BUL
##############################################################################

def ocr_resimleri_bul(
    soup,
    base_url
):

    adaylar = []

    for img in soup.find_all(
        "img"
    ):

        src = (
            img.get(
                "src"
            )
            or img.get(
                "data-src"
            )
            or img.get(
                "data-lazy-src"
            )
        )

        if not src:
            continue

        bilgi = (
            str(
                img.get(
                    "alt",
                    ""
                )
            )
            + " "
            + str(
                img.get(
                    "title",
                    ""
                )
            )
            + " "
            + str(
                img.get(
                    "class",
                    ""
                )
            )
            + " "
            + str(src)
        ).lower()

        skor = 0

        for keyword in OCR_KEYWORDS:

            if keyword in bilgi:
                skor += 10

        # Dosya isminde iletişim kelimesi varsa çok değerli
        if any(
            x in str(src).lower()
            for x in [
                "iletisim",
                "contact",
                "mail",
                "email"
            ]
        ):

            skor += 20

        # Küçük ikonları ele
        width = img.get(
            "width"
        )

        height = img.get(
            "height"
        )

        try:

            if width and int(
                str(width).replace(
                    "px",
                    ""
                )
            ) < MIN_IMAGE_WIDTH:

                continue

            if height and int(
                str(height).replace(
                    "px",
                    ""
                )
            ) < MIN_IMAGE_HEIGHT:

                continue

        except Exception:
            pass

        if skor <= 0:
            continue

        link = urljoin(
            base_url,
            src
        )

        adaylar.append(
            (
                skor,
                link
            )
        )

    adaylar.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return [
        link
        for _, link in adaylar[
            :MAX_OCR_IMAGE
        ]
    ]


##############################################################################
# OCR
##############################################################################

def resimden_mail_bul(
    image_url,
    mail_kaynaklari
):

    if not OCR_AKTIF:
        return set()

    if not OCR_VAR:
        return set()

    bulunan = set()

    try:

        r = session.get(
            image_url,
            timeout=(2, 5)
        )

        if r.status_code != 200:
            return set()

        content_type = r.headers.get(
            "Content-Type",
            ""
        ).lower()

        if (
            "image"
            not in content_type
        ):
            return set()

        image = Image.open(
            BytesIO(
                r.content
            )
        )

        width, height = image.size

        if (
            width < MIN_IMAGE_WIDTH
            or height < MIN_IMAGE_HEIGHT
        ):
            return set()

        ######################################################################
        # OCR
        ######################################################################

        text = pytesseract.image_to_string(
            image,
            lang="eng"
        )

        if not text:
            return set()

        text = ocr_mail_normalize(
            text
        )

        mailler = EMAIL_PATTERN.findall(
            text
        )

        for mail in mailler:

            mail = mail_normalize(
                mail
            )

            if mail:

                bulunan.add(
                    mail
                )

                mail_kaynaklari[
                    mail
                ] = "ocr"

    except Exception:
        pass

    return bulunan


##############################################################################
# SAYFA İÇİN OCR
##############################################################################

def sayfa_ocr(
    response,
    mail_kaynaklari,
    baslangic
):

    bulunan = set()

    if not OCR_AKTIF:
        return bulunan

    if not OCR_VAR:
        return bulunan

    try:

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        resimler = ocr_resimleri_bul(
            soup,
            response.url
        )

        for image_url in resimler:

            if (
                time.monotonic()
                - baslangic
                > MAX_FIRMA_SURE
            ):
                break

            mailler = resimden_mail_bul(
                image_url,
                mail_kaynaklari
            )

            bulunan.update(
                mailler
            )

    except Exception:
        pass

    return bulunan


##############################################################################
# TEK SAYFA TARA
##############################################################################

def sayfa_tara(
    url,
    mail_kaynaklari,
    baslangic,
    kaynak="html"
):

    bulunan = set()

    try:

        if (
            time.monotonic()
            - baslangic
            > MAX_FIRMA_SURE
        ):
            return bulunan, None

        response = session.get(
            url,
            timeout=(3, 6),
            allow_redirects=True
        )

        if response.status_code != 200:
            return bulunan, None

        content_type = response.headers.get(
            "Content-Type",
            ""
        ).lower()

        if (
            "html"
            not in content_type
            and "text"
            not in content_type
        ):
            return bulunan, None

        ######################################################################
        # HTML
        ######################################################################

        bulunan.update(
            sayfadan_mail_topla(
                response,
                mail_kaynaklari,
                kaynak
            )
        )

        ######################################################################
        # OCR
        ######################################################################

        bulunan.update(
            sayfa_ocr(
                response,
                mail_kaynaklari,
                baslangic
            )
        )

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        return bulunan, soup

    except Exception:

        return bulunan, None


##############################################################################
# ANA MAIL BULMA
##############################################################################

def mail_bul(site):

    baslangic = time.monotonic()

    bulunan = set()

    ziyaret = set()

    mail_kaynaklari = {}

    urller = olasi_urller(
        site
    )

    if not urller:
        return []

    ##########################################################################
    # DOMAIN
    ##########################################################################

    site_domain = url_temizle(
        site
    )

    if site_domain.startswith(
        "www."
    ):

        site_domain = site_domain[4:]

    ##########################################################################
    # ANA SİTELER
    ##########################################################################

    ana_soup = None
    ana_response_url = None

    for ana_url in urller:

        if (
            time.monotonic()
            - baslangic
            > MAX_FIRMA_SURE
        ):
            break

        if not gecerli_url(
            ana_url
        ):
            continue

        if ana_url in ziyaret:
            continue

        ziyaret.add(
            ana_url
        )

        bulunan_sayfa, soup = sayfa_tara(
            ana_url,
            mail_kaynaklari,
            baslangic,
            "html"
        )

        bulunan.update(
            bulunan_sayfa
        )

        if soup is not None:

            ana_soup = soup

            # Gerçek yönlendirilmiş URL
            try:

                ana_response_url = (
                    soup
                    and ana_url
                )

            except Exception:
                pass

            # İlk çalışan site yeterli
            break

    ##########################################################################
    # ANA SAYFADAN LİNKLERİ TOPLA
    ##########################################################################

    if ana_soup is not None:

        contact_links = (
            iletisim_linklerini_bul(
                ana_soup,
                ana_url
            )
        )

        ######################################################################
        # İletişim sayfalarını tara
        ######################################################################

        sayac = 0

        for link in contact_links:

            if (
                sayac
                >= MAX_CONTACT_PAGE
            ):
                break

            if (
                time.monotonic()
                - baslangic
                > MAX_FIRMA_SURE
            ):
                break

            if link in ziyaret:
                continue

            ziyaret.add(
                link
            )

            sayac += 1

            sayfa_mailler, _ = (
                sayfa_tara(
                    link,
                    mail_kaynaklari,
                    baslangic,
                    "contact"
                )
            )

            bulunan.update(
                sayfa_mailler
            )

    ##########################################################################
    # SON FİLTRE
    ##########################################################################

    temiz_mailler = temizle(
        list(bulunan)
    )

    ##########################################################################
    # LOG
    ##########################################################################

    if temiz_mailler:

        en_iyi = en_iyi_mail_sec(
            temiz_mailler,
            site_domain,
            mail_kaynaklari
        )

        print(
            f"  MAIL ADAYLARI: "
            f"{len(temiz_mailler)} -> "
            f"{en_iyi}"
        )

    return temiz_mailler


##############################################################################
# EXCEL OKU
##############################################################################

print()
print("=" * 70)
print("Excel okunuyor...")
print("=" * 70)

df = pd.read_excel(
    INPUT
)

print(
    f"Toplam firma: {len(df)}"
)


##############################################################################
# GEREKLİ KOLONLAR
##############################################################################

# Sadece WEB sütunu zorunlu.
# Büyük/küçük harf fark etmez:
# WEB / web / Web / wEb / WeB ... hepsi kabul edilir.

kolon_map = {
    str(kolon).strip().lower(): kolon
    for kolon in df.columns
}

if "web" not in kolon_map:

    raise ValueError(
        "Excel'de zorunlu olan 'WEB' sütunu bulunamadı.\n"
        "WEB / web / Web gibi yazımların tamamı kabul edilir."
    )

WEB_KOLONU = kolon_map["web"]
UNVAN_KOLONU = kolon_map.get("unvan")
SICIL_KOLONU = kolon_map.get("sicil")

##############################################################################
# TEK FİRMA İŞLE
##############################################################################

def firma_isle(satir):

    # UNVAN varsa kullan
    if UNVAN_KOLONU is not None:

        firma = str(
            satir[UNVAN_KOLONU]
        ).strip()

        if firma.lower() == "nan":
            firma = ""

    else:

        firma = ""

    # SİCİL varsa kullan
    if SICIL_KOLONU is not None:

        sicil = satir[SICIL_KOLONU]

    else:

        sicil = ""

    # WEB sütunu zorunlu
    site_deger = satir[WEB_KOLONU]

    if pd.isna(site_deger):

        site = ""

    else:

        site = str(
            site_deger
        ).strip()

    ##########################################################################
    # WEB YOK
    ##########################################################################

    if (
        site == ""
        or site.lower() == "nan"
    ):

        return {
            "SİCİL": sicil,
            "UNVAN": firma,
            "WEB": "",
            "EMAIL": ""
        }

    ##########################################################################
    # MAIL ARA
    ##########################################################################

    try:

        adaylar = mail_bul(
            site
        )

        temiz_site = url_temizle(
            site
        )

        if temiz_site.startswith(
            "www."
        ):

            site_domain = temiz_site[4:]

        else:

            site_domain = temiz_site

        mail = en_iyi_mail_sec(
            adaylar,
            site_domain
        )

    except Exception as e:

        print(
            f"HATA: {firma or site} -> {e}"
        )

        mail = ""

    ##########################################################################
    # SONUÇ
    ##########################################################################

    return {
        "SİCİL": sicil,
        "UNVAN": firma,
        "WEB": site,
        "EMAIL": mail
    }


##############################################################################
# PARALEL ÇALIŞTIR
##############################################################################

print()
print("=" * 70)
print("Mail adresleri aranıyor...")
print("=" * 70)
print()

sonuclar = [
    None
] * len(df)


with ThreadPoolExecutor(
    max_workers=MAX_WORKERS
) as executor:

    futures = {
        executor.submit(
            firma_isle,
            satir
        ): i

        for i, (_, satir)
        in enumerate(
            df.iterrows(),
            start=1
        )
    }

    toplam = len(futures)

    for sira, future in enumerate(
        as_completed(futures),
        start=1
    ):

        try:

            sonuc = future.result()

        except Exception as e:

            print(
                f"HATA: {e}"
            )

            continue

        index = futures[
            future
        ]

        sonuclar[
            index - 1
        ] = sonuc

        ######################################################################
        # EKRAN
        ######################################################################

        if sonuc[
            "EMAIL"
        ]:

            print(
                f"{sira}/{toplam} ✓ "
                f"{sonuc['UNVAN']} -> "
                f"{sonuc['EMAIL']}"
            )

        else:

            print(
                f"{sira}/{toplam} - "
                f"{sonuc['UNVAN']}"
            )


##############################################################################
# BOŞ SONUÇLARI TEMİZLE
##############################################################################

sonuclar = [
    x
    for x in sonuclar
    if x is not None
]


##############################################################################
# DATAFRAME
##############################################################################

sonuc_df = pd.DataFrame(
    sonuclar
)


##############################################################################
# EXCEL FORMÜL KONTROL
##############################################################################

for kolon in [
    "WEB",
    "EMAIL"
]:

    if kolon in sonuc_df.columns:

        sonuc_df[kolon] = (
            sonuc_df[kolon]
            .apply(excel_text)
        )


##############################################################################
# EXCEL KAYDET
##############################################################################

sonuc_df.to_excel(
    OUTPUT,
    index=False
)


##############################################################################
# İSTATİSTİK
##############################################################################

bulunan = (

    sonuc_df["EMAIL"]
    .fillna("")
    .astype(str)
    .str.strip()
    .ne("")
    .sum()
)


print()
print("=" * 70)
print("TAMAMLANDI")
print("=" * 70)

print(
    f"Toplam Firma : {len(sonuc_df)}"
)

print(
    f"Mail Bulunan : {bulunan}"
)

if len(sonuc_df):

    print(
        f"Başarı Oranı : "
        f"%{bulunan / len(sonuc_df) * 100:.1f}"
    )

print()
print(
    f"Dosya: {OUTPUT}"
)

if not OCR_VAR:

    print()
    print(
        "UYARI: OCR aktif değil. "
        "pytesseract ve Pillow kurulu değil."
    )
