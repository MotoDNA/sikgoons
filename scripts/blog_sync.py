#!/usr/bin/env python3
"""네이버 블로그(ff9302) 글을 가져와 sikgoons.kr/blog/ 정적 페이지로 만든다.

GitHub Actions에서 3시간마다 실행된다.
- RSS로 최근 글 목록을 읽고, 새 글이면 본문을 가져와 blog/posts.json 에 쌓는다.
- blog/index.html(목록)과 blog/<글번호>/index.html(본문)을 다시 만든다.
- sitemap.xml 에 블로그 주소를 더한다.
원본 글은 네이버 블로그에 그대로 있고, 각 글 페이지의 canonical 은 원문을 가리킨다.
"""
import html, json, os, re, sys, time, urllib.request
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime
from bs4 import BeautifulSoup

BLOG_ID = 'ff9302'
SITE = 'https://sikgoons.kr'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BLOG_DIR = os.path.join(ROOT, 'blog')
DATA = os.path.join(BLOG_DIR, 'posts.json')
TPL = os.path.join(ROOT, '_blog', 'template.html')
UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'
STATIC_ROUTES = ['', 'works/', 'about/', 'contact/', 'blog/']


def get(url):
    req = urllib.request.Request(url, headers={'User-Agent': UA, 'Referer': f'https://blog.naver.com/{BLOG_ID}'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', 'replace')


def img_url(src):
    if not src:
        return None
    src = src.split('?')[0]
    if 'pstatic.net' not in src and 'naver.net' not in src:
        return None
    return src + '?type=w966'


def parse_post(log_no):
    """본문을 블록 목록으로 바꾼다: {t:'p'|'h'|'q'|'img'|'hr', ...}"""
    page = get(f'https://blog.naver.com/PostView.naver?blogId={BLOG_ID}&logNo={log_no}&redirect=Dlog&widgetTypeCall=true&directAccess=false')
    soup = BeautifulSoup(page, 'html.parser')
    blocks = []
    main = soup.select_one('.se-main-container')
    if main:  # 스마트에디터 ONE
        for comp in main.select(':scope > .se-component'):
            cls = ' '.join(comp.get('class', []))
            if 'se-image' in cls or 'se-imageGroup' in cls or 'se-imageStrip' in cls:
                for im in comp.select('img'):
                    u = img_url(im.get('data-lazy-src') or im.get('src'))
                    if u:
                        blocks.append({'t': 'img', 'src': u})
                cap = comp.select_one('.se-caption')
                if cap and cap.get_text(strip=True):
                    blocks.append({'t': 'cap', 'text': cap.get_text(' ', strip=True)})
            elif 'se-sectionTitle' in cls:
                tx = comp.get_text(' ', strip=True)
                if tx:
                    blocks.append({'t': 'h', 'text': tx})
            elif 'se-quotation' in cls:
                tx = comp.get_text(' ', strip=True)
                if tx:
                    blocks.append({'t': 'q', 'text': tx})
            elif 'se-horizontalLine' in cls:
                blocks.append({'t': 'hr'})
            elif 'se-text' in cls:
                for p in comp.select('.se-text-paragraph'):
                    tx = p.get_text('', strip=False).replace('​', '').strip()
                    blocks.append({'t': 'p', 'text': tx})
    else:  # 예전 에디터
        old = soup.select_one('#postViewArea') or soup.select_one('.post-view')
        if old:
            for el in old.find_all(['p', 'img', 'div'], recursive=True):
                if el.name == 'img':
                    u = img_url(el.get('data-lazy-src') or el.get('src'))
                    if u:
                        blocks.append({'t': 'img', 'src': u})
                elif el.name == 'p':
                    tx = el.get_text(' ', strip=True)
                    if tx:
                        blocks.append({'t': 'p', 'text': tx})
    # 빈 줄 여러 개를 하나로
    out = []
    for b in blocks:
        if b['t'] == 'p' and not b['text'] and out and out[-1]['t'] == 'p' and not out[-1]['text']:
            continue
        out.append(b)
    while out and out[0]['t'] == 'p' and not out[0]['text']:
        out.pop(0)
    return out


def load():
    if os.path.exists(DATA):
        return json.load(open(DATA, encoding='utf-8'))
    return []


def sync():
    posts = {p['id']: p for p in load()}
    rss = get(f'https://rss.blog.naver.com/{BLOG_ID}.xml')
    root = ET.fromstring(rss)
    new = 0
    for it in root.iter('item'):
        link = (it.findtext('link') or '').strip()
        m = re.search(r'/(\d{6,})', link)
        if not m:
            continue
        pid = m.group(1)
        title = html.unescape((it.findtext('title') or '').strip())
        date = parsedate_to_datetime(it.findtext('pubDate')).strftime('%Y-%m-%d')
        cat = (it.findtext('category') or '').strip()
        desc_html = it.findtext('description') or ''
        dsoup = BeautifulSoup(desc_html, 'html.parser')
        summary = re.sub(r'\s+', ' ', dsoup.get_text(' ', strip=True))[:160]
        thumb = None
        im = dsoup.find('img')
        if im:
            thumb = img_url(im.get('src'))
        if pid in posts and posts[pid].get('blocks'):
            posts[pid].update(title=title, cat=cat)
            continue
        try:
            blocks = parse_post(pid)
            time.sleep(1.5)
        except Exception as e:  # 본문을 못 읽으면 요약만 보여 준다
            print('본문 실패', pid, e, file=sys.stderr)
            blocks = []
        if not thumb:
            thumb = next((b['src'] for b in blocks if b['t'] == 'img'), None)
        posts[pid] = dict(id=pid, title=title, date=date, cat=cat, summary=summary, thumb=thumb,
                          url=f'https://blog.naver.com/{BLOG_ID}/{pid}', blocks=blocks)
        new += 1
    lst = sorted(posts.values(), key=lambda p: (p['date'], p['id']), reverse=True)
    os.makedirs(BLOG_DIR, exist_ok=True)
    json.dump(lst, open(DATA, 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
    print(f'글 {len(lst)}개 (새 글 {new}개)')
    return lst


def esc(s):
    return html.escape(s or '', quote=True)


def page(tpl, title, desc, path, body, canonical=None, og_image=None):
    canonical = canonical or f'{SITE}/{path}'
    head = (f'<title>{esc(title)}</title>\n<meta name="description" content="{esc(desc)}">\n'
            f'<link rel="canonical" href="{esc(canonical)}">\n<meta property="og:type" content="article">\n'
            f'<meta property="og:title" content="{esc(title)}">\n<meta property="og:description" content="{esc(desc)}">\n'
            f'<meta property="og:url" content="{SITE}/{path}">\n<meta property="og:image" content="{esc(og_image or SITE + "/img/d6027.jpg")}">')
    return tpl.replace('<!--HEAD-->', head).replace('<!--BODY-->', body)


def render(lst):
    tpl = open(TPL, encoding='utf-8').read()
    cards = []
    for p in lst:
        th = (f'<img src="{esc(p["thumb"])}" alt="" loading="lazy" referrerpolicy="no-referrer">' if p.get('thumb') else '<span class="bthumb-empty">SIKGOONS</span>')
        cards.append(f'<a class="bcard" href="/blog/{p["id"]}/"><span class="bthumb">{th}</span>'
                     f'<span class="bmeta">{p["date"].replace("-", ".")}{(" · " + esc(p["cat"])) if p.get("cat") else ""}</span>'
                     f'<span class="btitle">{esc(p["title"])}</span><span class="bsum">{esc(p.get("summary", ""))}</span></a>')
    body = ('<section class="wrap sec" style="padding-top:48px"><div class="eyebrow">Blog</div>'
            '<h1 class="h2" style="font-size:clamp(30px,3.4vw,48px)">시공 일지</h1>'
            '<p class="note" style="margin:-10px 0 30px">네이버 블로그에 올린 글이 이곳에 자동으로 모입니다.</p>'
            f'<div class="bgrid">{"".join(cards)}</div>'
            f'<p style="margin-top:40px"><a class="btn" href="https://blog.naver.com/{BLOG_ID}" target="_blank" rel="noopener">네이버 블로그에서 전체 글 보기 ↗</a></p></section>')
    open(os.path.join(BLOG_DIR, 'index.html'), 'w', encoding='utf-8').write(
        page(tpl, '시공 일지 | 식군스인테리어', '식군스인테리어가 네이버 블로그에 기록한 양산·부산 아파트 인테리어 시공 일지.', 'blog/', body))
    for i, p in enumerate(lst):
        parts = []
        for b in p.get('blocks') or []:
            t = b['t']
            if t == 'p':
                parts.append(f'<p>{esc(b["text"])}</p>' if b['text'] else '<p class="gap"></p>')
            elif t == 'h':
                parts.append(f'<h2>{esc(b["text"])}</h2>')
            elif t == 'q':
                parts.append(f'<blockquote>{esc(b["text"])}</blockquote>')
            elif t == 'cap':
                parts.append(f'<p class="cap">{esc(b["text"])}</p>')
            elif t == 'hr':
                parts.append('<hr>')
            elif t == 'img':
                parts.append(f'<img src="{esc(b["src"])}" alt="" loading="lazy" referrerpolicy="no-referrer">')
        if not parts:
            parts.append(f'<p>{esc(p.get("summary", ""))}</p>')
        prev = lst[i + 1] if i + 1 < len(lst) else None
        nxt = lst[i - 1] if i > 0 else None
        pager = '<nav class="pager">' + (f'<a href="/blog/{prev["id"]}/"><small>← 이전 글</small>{esc(prev["title"])}</a>' if prev else '<span></span>') + \
                '<a href="/blog/" style="align-items:center"><small>LIST</small>목록</a>' + \
                (f'<a href="/blog/{nxt["id"]}/" style="text-align:right"><small>다음 글 →</small>{esc(nxt["title"])}</a>' if nxt else '<span></span>') + '</nav>'
        body = (f'<article class="wrap sec post" style="padding-top:48px"><div class="eyebrow">{p["date"].replace("-", ".")}{(" · " + esc(p["cat"])) if p.get("cat") else ""}</div>'
                f'<h1>{esc(p["title"])}</h1><div class="pbody">{"".join(parts)}</div>'
                f'<p class="orig"><a class="btn" href="{esc(p["url"])}" target="_blank" rel="noopener">네이버 블로그 원문 보기 ↗</a> '
                f'<a class="btn" href="/contact/">상담 신청하기 →</a></p>{pager}</article>')
        d = os.path.join(BLOG_DIR, p['id'])
        os.makedirs(d, exist_ok=True)
        open(os.path.join(d, 'index.html'), 'w', encoding='utf-8').write(
            page(tpl, f'{p["title"]} | 식군스인테리어', p.get('summary') or p['title'], f'blog/{p["id"]}/', body,
                 canonical=p['url'], og_image=p.get('thumb')))
    # 사이트맵
    works = sorted(n for n in os.listdir(os.path.join(ROOT, 'works')) if os.path.isdir(os.path.join(ROOT, 'works', n)))
    urls = STATIC_ROUTES + [f'works/{w}/' for w in works]
    sm = ['<?xml version="1.0" encoding="UTF-8"?>', '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">']
    sm += [f'  <url><loc>{SITE}/{u}</loc></url>' for u in urls]
    # 글 페이지는 원문(네이버)을 canonical 로 두므로 사이트맵에는 목록 페이지만 넣는다
    sm.append('</urlset>')
    open(os.path.join(ROOT, 'sitemap.xml'), 'w', encoding='utf-8').write('\n'.join(sm) + '\n')


if __name__ == '__main__':
    if '--render-only' in sys.argv:
        render(load())
    else:
        render(sync())
