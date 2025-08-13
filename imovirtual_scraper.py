#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Imovirtual → CSV + PowerPoint
Autor: preparado para Hugo Silva (RE/MAX Oceanus)

AVISO LEGAL: Usa apenas dados/fotos com autorização. O scraping pode violar Termos de Uso.
"""

import asyncio
import argparse
import csv
import io
import json
import re
import time
from pathlib import Path
from typing import Dict, Any, List, Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

# ---------------------------
# Navegação com Playwright
# ---------------------------
async def fetch_html(url: str, render: str = "networkidle", timeout_ms: int = 45000) -> str:
    """
    Abre a página com Chromium headless e devolve o HTML final.
    render: "load" | "domcontentloaded" | "networkidle"
    """
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
        ))
        page = await context.new_page()
        await page.goto(url, wait_until=render, timeout=timeout_ms)
        html = await page.content()
        await browser.close()
        return html

# ---------------------------
# Parsing auxiliar
# ---------------------------
def text_or_none(el) -> Optional[str]:
    return el.get_text(strip=True) if el else None

def find_label_value(soup: BeautifulSoup, labels: List[str]) -> Optional[str]:
    """
    Procura pares 'label: valor' em detalhes do anúncio.
    Tenta padrões comuns (dt/dd, li, div) e matching por prefixo (case-insensitive).
    """
    label_pattern = re.compile(r"^\s*(%s)\s*:?$" % "|".join([re.escape(x) for x in labels]), re.I)

    # Padrão <dt>Label</dt><dd>Valor</dd>
    for dt in soup.find_all("dt"):
        if dt.string and label_pattern.match(dt.get_text(strip=True)):
            dd = dt.find_next("dd")
            if dd:
                return dd.get_text(" ", strip=True)

    # Padrão listas
    for li in soup.find_all("li"):
        strong = li.find(["strong", "span"])
        if strong and label_pattern.match(strong.get_text(strip=True)):
            txt = li.get_text(" ", strip=True)
            lab = strong.get_text(strip=True)
            val = txt.replace(lab, "").strip(" :\u00a0-")
            if val:
                return val

    # Padrão divs genéricos
    for node in soup.find_all(["div", "span"]):
        if node.get_text(strip=True) and label_pattern.match(node.get_text(strip=True)):
            sib = node.find_next_sibling()
            if sib and sib.get_text(strip=True):
                return sib.get_text(" ", strip=True)

    return None

def parse_json_ld(soup: BeautifulSoup) -> Dict[str, Any]:
    """
    Extrai JSON-LD relevante (RealEstate, Offer, etc.).
    """
    data: Dict[str, Any] = {}
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(tag.string or "{}")
        except Exception:
            continue
        # Pode vir como lista
        if isinstance(obj, list):
            for x in obj:
                _merge_realestate(data, x)
        else:
            _merge_realestate(data, obj)
    return data

def _merge_realestate(dst: Dict[str, Any], obj: Dict[str, Any]):
    if not isinstance(obj, dict):
        return
    if "@graph" in obj and isinstance(obj["@graph"], list):
        for g in obj["@graph"]:
            _merge_realestate(dst, g)

    offer = obj.get("offers") or {}
    address = obj.get("address") or {}

    dst.setdefault("title", obj.get("name"))
    dst.setdefault("description", obj.get("description"))

    if isinstance(offer, dict):
        price = offer.get("price") or offer.get("lowPrice")
        dst.setdefault("price", price)
        currency = offer.get("priceCurrency")
        if price and currency:
            dst["price_str"] = f"{price} {currency}"

    if isinstance(address, dict):
        loc = " ".join([
            str(address.get("addressLocality") or ""),
            str(address.get("addressRegion") or ""),
            str(address.get("addressCountry") or ""),
        ]).strip()
        if loc:
            dst.setdefault("location", loc)

    rooms = obj.get("numberOfRooms") or obj.get("numberOfBedrooms")
    if rooms:
        dst.setdefault("bedrooms", str(rooms))

# ---------------------------
# Parser principal
# ---------------------------
def parse_listing(html: str, url: str, max_images: int = 3) -> Dict[str, Any]:
    soup = BeautifulSoup(html, "lxml")

    # 1) Tenta JSON-LD primeiro (mais estável)
    meta = parse_json_ld(soup)

    # 2) Título
    title = meta.get("title") or text_or_none(soup.find("h1"))

    # 3) Preço
    price = meta.get("price")
    price_str = meta.get("price_str")
    if not price_str:
        price_el = soup.find(["strong", "span"], attrs={"aria-label": re.compile("Preço", re.I)})
        price_str = text_or_none(price_el) or (str(price) if price else "")

    # 4) Localização
    location = meta.get("location")
    if not location:
        # fallback genérico
        breadcrumb = soup.select_one("nav[aria-label='breadcrumb']")
        location = text_or_none(breadcrumb) or ""

    # 5) Tipologia / Quartos / WC / Área
    typology   = find_label_value(soup, ["Tipologia"])
    bedrooms   = meta.get("bedrooms") or find_label_value(soup, ["Quartos", "Nº de quartos", "Número de quartos"])
    bathrooms  = find_label_value(soup, ["Casas de banho", "WCs"])
    area       = find_label_value(soup, ["Área bruta", "Área útil", "Área", "Área (m²)", "Área bruta (m²)"])

    # 6) Descrição
    description = meta.get("description")
    if not description:
        paragraphs = [p.get_text(" ", strip=True) for p in soup.find_all("p")]
        paragraphs.sort(key=len, reverse=True)
        description = paragraphs[0] if paragraphs else ""

    # 7) Imagens
    images: List[str] = []
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or img.get("data-lazy")
        if not src:
            continue
        if src.startswith("//"):
            src = "https:" + src
        if src.startswith("http"):
            images.append(src)
        if len(images) >= max_images:
            break

    rec: Dict[str, Any] = {
        "url": url,
        "title": (title or "").strip(),
        "price": (price_str or "").strip(),
        "location": (location or "").strip(),
        "area": (area or "").strip(),
        "typology": (typology or "").strip(),
        "bedrooms": (str(bedrooms) if bedrooms else "").strip(),
        "bathrooms": (str(bathrooms) if bathrooms else "").strip(),
        "description": (description or "").strip(),
    }
    for i, src in enumerate(images, start=1):
        rec[f"image{i}"] = src
    return rec

# ---------------------------
# PowerPoint
# ---------------------------
def add_titlebox(slide, left, top, width, height, text, font_size=28, bold=T
