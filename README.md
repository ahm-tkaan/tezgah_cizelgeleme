# Üretim Çizelgeleme Sistemi (Production Scheduling System)

Bu sistem, üretim süreçlerinde iş emirlerini makinelere atamak için geliştirilmiş bir çizelgeleme algoritmasıdır. Kesici uç kullanımını optimize ederek iş yükünü dengeli bir şekilde dağıtmayı amaçlar.

## Özellikler

- İş emirlerini önem derecesine göre sıralama
- Kesici uç kullanımına dayalı otomatik makine atama
- Makine yüklerini dengeleme
- İş emri önceliklerini dikkate alma
- Detaylı çizelge ve özet raporları oluşturma

## Gereksinimler

- Python 3.6+
- Pandas
- NumPy
- Openpyxl

```bash
pip install pandas numpy openpyxl
```

## Veri Dosyaları

Uygulamanın çalışması için aşağıdaki Excel dosyalarının `veriler` klasöründe bulunması gerekmektedir:

- **Ürün Havuzu.xlsx**: İş emirlerini içeren ana dosya
- **Çizelgeleme Önem Sırası.xlsx**: İş emirlerinin önem sırasını belirleyen dosya
- **Standart süre.xlsx**: Ürünlerin standart üretim sürelerini içeren dosya
- **Ürün kesici uç.xlsx**: Ürünlerin kesici uç bilgilerini içeren dosya

## Kullanım

Script'i doğrudan çalıştırabilirsiniz:

```bash
python uretim_cizelgeleme.py
```

Çalıştırıldığında, script otomatik olarak verileri işleyecek ve sonuçları `uretim_cizelgeleme_sonuclari.xlsx` dosyasına kaydedecektir.

## Çalışma Mantığı

1. **Veri Yükleme ve Hazırlama**: Tüm veri kaynaklarından bilgiler yüklenir ve birleştirilir
2. **Makine Atama Kuralları Oluşturma**: Kesici uç kullanım oranlarına göre makine atama kuralları belirlenir
3. **Yüksek Öncelikli İşleri Atama**: Önem skoru 0.2 ve üzerinde olan işler öncelikli olarak makinelere atanır
4. **Düşük Öncelikli İşleri Atama**: Kalan işler makine yüklerini dengeleyecek şekilde atanır
5. **Sonuçları Raporlama**: Excel dosyasına detaylı çizelge ve özet bilgiler kaydedilir

## Çıktı Dosyası

Oluşturulan Excel dosyası aşağıdaki sayfaları içerir:

- **Çizelgeleme Sonuçları**: Tüm iş emirlerinin atandığı makineler ve süre bilgileri
- **Tezgah Özeti**: Her makinenin toplam iş yükü ve atanan iş sayısı
- **Kesici Uç Özeti**: Kesici uçların kullanım sıklığı ve bulundukları makineler
- **Tezgah Zaman Çizelgesi**: Her makinenin zaman içindeki iş yükü dağılımı

## Ana Fonksiyonlar

- `load_and_prepare_data()`: Veri dosyalarını yükler ve hazırlar
- `create_machine_assignment_rules()`: Kesici uç kullanım istatistiklerine göre makine atama kurallarını oluşturur
- `assign_machines_high_importance()`: Yüksek öncelikli işleri makinelere atar
- `assign_machines_low_importance()`: Düşük öncelikli işleri makinelere atar
- `export_to_excel()`: Sonuçları Excel dosyasına kaydeder
- `main()`: Ana uygulama akışını yönetir
