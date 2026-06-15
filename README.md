# Chang bo'roni va havo sifati prognoz tizimi

O'zbekiston viloyatlari uchun chang bo'ronlarini va Toshkent shahri uchun havo ifloslanishini prognoz qiladigan model.

## Loyiha tuzilishi

```
dust_storm_all/
├── analysis/                    # Tahlil skriptlari
│   ├── step0_data_quality.py   # 0-bosqich: ma'lumot sifati va chang bo'roni aniqlash
│   └── results/                # Tahlil natijalari (CSV, TXT)
├── data/                       # Ma'lumotlar (gitignore, alohida yuklanadi)
│   └── все станции за 10 лет/  # 84 stansiya, 2011-2020
└── все станции за 10 лет.zip   # Asl ma'lumot arxivi
```

## Ma'lumotlar

- **Manba:** O'zgidromet (O'zbekiston Gidrometeorologiya xizmati)
- **Davr:** 2011-2020 (10 yil), ba'zi stansiyalar 1961-dan
- **Chastota:** Kunlik
- **Stansiyalar:** 84 ta (barcha viloyatlar)
- **Parametrlar:** 38 ta (harorat, namlik, shamol, ko'rinish, bosim, yog'in, tuproq va h.k.)

### Asosiy parametrlar

| Kod | Tavsif | Birlik |
|-----|--------|--------|
| V | Ko'rinish (visibility) | km |
| Vx8 | Shamol tezligi | m/s |
| VxG | Shamol loli (gust) | m/s |
| U / UN | Nisbiy namlik (o'rtacha / minimal) | % |
| Taav | O'rtacha havo harorati | °C |
| StP / SeP | Stansiya / dengiz sathi bosimi | hPa |

## 0-Bosqich natijalari

Tuman filtrlangan chang bo'roni aniqlash algoritmi natijalari:

- **Jami hodisalar:** 71,582
- **Kuchli (V < 0.5 km + shamol + quruqlik):** 12,663
- **O'rtacha (V < 1 km + shamol/quruqlik):** 35,851
- **Yengil (V < 2 km + shamol + quruqlik):** 23,068

### Eng faol stansiyalar (Top-5)
1. NASREDINBEK — 2,642 hodisa
2. Pap — 3,062 hodisa
3. Nurabod — 3,139 hodisa
4. KOKARAL — 2,664 hodisa
5. SARKANDA — 3,044 hodisa

### Mavsumiylik
Eng faol oylar: Sentabr-Noyabr (kuz), eng past: Fevral-Aprel

## Ishga tushirish

```bash
# xlrd kutubxonasi kerak (XLS fayllarni o'qish uchun)
pip install xlrd

# 0-bosqich tahlilini ishga tushirish
python analysis/step0_data_quality.py
```

## Roadmap

- [x] 0-bosqich: Ma'lumot sifati tahlili
- [ ] 1-bosqich: Chang bo'roni hodisalari jadvalini yakunlash
- [ ] 2-bosqich: ERA5 + Sentinel-5P ma'lumotlarini ulash
- [ ] 3-bosqich: ML model (XGBoost/LightGBM)
- [ ] 4-bosqich: Toshkent havo sifati modeli
- [ ] 5-bosqich: O'zbekiston xaritasi vizualizatsiyasi
- [ ] 6-bosqich: Telegram bot
