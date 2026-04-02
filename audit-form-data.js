/**
 * Pizzabulls Operasyon Denetim Ziyaret Formu — soru listesi ve maksimum puanlar.
 * Kalite 46, Servis 56, Temizlik ve Güvenlik 40 (toplam 142).
 */
(function attachOperationAuditForm() {
  const OPERATION_AUDIT_FORM = {
    title: "PIZZABULLS OPERASYON DENETİM ZİYARET FORMU",
    gradeScaleNote:
      "Skala: A %93–100, B %85–92.9, C %75–84.9, F %74.9 ve altı.",
    sections: [
      {
        key: "kalite",
        title: "KALİTE",
        maxTotal: 46,
        items: [
          { text: "Onaylı ürünler kullanılıyor mu?", max: 2 },
          { text: "Çiğ ürünlerin durumu", max: 2 },
          { text: "Hamurların SKT'leri", max: 2 },
          { text: "Soğuk Odadaki Açık Ürünlerin SKT'leri", max: 2 },
          { text: "Ürünlerin rotasyonu", max: 2 },
          {
            text: "Yan Ürünlerin Pişirme prosedürü (Soğan Halkası ve Patbulls Donuk pişiriliyor.)",
            max: 2,
          },
          { text: "Macline'daki ürünlerin SKT'leri", max: 2 },
          {
            text: "Alt, Üst Mozzarella kullanımı standartlara uyuyor mu?",
            max: 2,
          },
          { text: "Pizza kenarları standart", max: 2 },
          {
            text: "Ürün Dağılımları standartlara uygun mu?",
            max: 2,
          },
          { text: "Ürünler yerle temas ediyor mu?", max: 2 },
          {
            text: "Derin dondurucuda tüm ürünler kapalı mı?",
            max: 2,
          },
          { text: "Stok odası temiz düzenli mi?", max: 2 },
          { text: "Soğuk oda temiz düzenli mi?", max: 2 },
          {
            text: "Pizza sosu hazırlama ve saklama prosedürleri uygulanıyor mu?",
            max: 2,
          },
          {
            text: "Yan Ürünler Tartılarak porsiyonlanıyor mu?",
            max: 2,
          },
          {
            text: "Soğutucu & Dondurucu ekipman ısıları aralıkları uygun mu?",
            max: 2,
          },
          { text: "Deefrezde buzlanma var mı?", max: 2 },
          { text: "Yeterli sayıda Screen var mı?", max: 2 },
          { text: "Terazi Çalışıyor mu?", max: 2 },
          {
            text: "Önceki Ziyaret Raporuna Göre Harekete geçilmiş mi?",
            max: 2,
          },
          {
            text: "Mozarella çözündürme ve saklama prosedürleri uygulanıyor mu?",
            max: 2,
          },
          {
            text: "Walk-in (Bakımlı mı perdeler tam, ışıklar korumalı ve yanıyor)",
            max: 2,
          },
        ],
      },
      {
        key: "servis",
        title: "SERVİS",
        maxTotal: 56,
        items: [
          {
            text: "Restoran kadrosu incelendiğinde kadro yeterli mi?",
            max: 2,
          },
          {
            text: "Alınan bütün siparişler sisteme kaydedilmekte mi?",
            max: 2,
          },
          {
            text: "Müdür giyim tarzı iş görsellerinde görüldüğü gibi",
            max: 2,
          },
          {
            text: "Ekip giyim tarzı iş görsellerinde görüldüğü gibi",
            max: 2,
          },
          {
            text: "Güncel menü ve el ilanları kullanılıyor mu?",
            max: 2,
          },
          { text: "Bütün sürücüler kask takıyor mu?", max: 2 },
          {
            text: "Servis standartlarına uyuluyor mu? Masa ve Gel-Al",
            max: 2,
          },
          {
            text: "Menüdeki bütün ürünler restoranda mevcut mu?",
            max: 2,
          },
          { text: "Şubede yeterli Panç var mı?", max: 2 },
          {
            text: "Anlık Satış Hacmine uygun ekip var mı?",
            max: 2,
          },
          { text: "Kutu kapama stickerları kullanılıyor mu?", max: 2 },
          {
            text: "Tüm iletişim mecraları çalışmakta mı?",
            max: 2,
          },
          {
            text: "Pizza paketleme standartlarına uyuluyor? (Pizza ayağı, Islak mendil, Pizza baharatı)",
            max: 2,
          },
          {
            text: "Sürücüler ürünleri eksiksiz götürüyor?",
            max: 2,
          },
          { text: "Rush Hazırlığı yapılmış mı?", max: 2 },
          {
            text: "Siparişler zamanında götürülüyor?",
            max: 2,
          },
          {
            text: "Ön banko siparişlerini zamanında teslim etmekte mi?",
            max: 2,
          },
          { text: "İş yardımcıları güncel mi?", max: 2 },
          {
            text: "Restoran içi pazarlama materyalleri uygun mu?",
            max: 2,
          },
          {
            text: "Restoran dışı pazarlama materyalleri uygun mu?",
            max: 2,
          },
          {
            text: "LCD Menü- Board çalışıyor, Menüler Güncel",
            max: 2,
          },
          {
            text: "Restoran aydınlatmaları çalışıyor mu?",
            max: 2,
          },
          { text: "Personellerin Hijyen belgeleri varmı.?", max: 1 },
          {
            text: "Motorlar ve Kasaları Temiz ve bakımlı mı?",
            max: 5,
          },
          { text: "Online Sipariş Puanı YS", max: 2 },
          { text: "Online Sipariş Puanı Gyi", max: 2 },
          { text: "Online Sipariş Puanı Tyi", max: 2 },
        ],
      },
      {
        key: "temizlik",
        title: "TEMİZLİK VE GÜVENLİK",
        maxTotal: 40,
        items: [
          {
            text: "Dış reklamlar temiz, bakımlı ve ışıkları tam mı?",
            max: 2,
          },
          { text: "Camlar temiz mi?", max: 2 },
          {
            text: "Masa sandalyeler temiz ve bakımlı mı?",
            max: 2,
          },
          {
            text: "Lobi yerler ve duvarlar temiz mi?",
            max: 2,
          },
          {
            text: "Mutfak yerler ve duvarlar temiz mi?",
            max: 2,
          },
          { text: "Pançlar temiz ve bakımlı mı?", max: 2 },
          {
            text: "Bezlerin kullanımı, ayrı bez kovaları var mı?",
            max: 2,
          },
          {
            text: "El yıkama prosedürlerine dikkat ediliyor mu?",
            max: 2,
          },
          {
            text: "Dış alan temiz bakımlı Pizzabulls Logolu çöp yok",
            max: 2,
          },
          {
            text: "Çöp kovası temiz ve bakımlı mı?",
            max: 2,
          },
          {
            text: "Davlumbaz temiz ve bakımlı mı?",
            max: 2,
          },
          { text: "Fırın temiz ve bakımlı mı?", max: 2 },
          { text: "Tuvaletler temiz mi?", max: 2 },
          {
            text: "Temizlik Materyalleri Tam mı? (Kağıt Havlu, El Sabunu vb.)",
            max: 2,
          },
          {
            text: "Periyodik ilaçlama yapılmakta mı?",
            max: 2,
          },
          {
            text: "Tüm ekipmanlar çalışıyor, arızalı ekipman yok",
            max: 2,
          },
          {
            text: "Yangın tüpü mevcut mu? Dolumları yapılmış mı?",
            max: 2,
          },
          { text: "Ecza Dolabı var mı?", max: 2 },
          {
            text: "Ecza dolabında gerekli malzemeler mevcut mu?",
            max: 2,
          },
          {
            text: "Çalışma İstasyonlarında risk var mı (Priz bozukluğu, açık kablo v.b)",
            max: 2,
          },
        ],
      },
    ],
  };

  /** @type {Window & { OPERATION_AUDIT_FORM: typeof OPERATION_AUDIT_FORM }} */
  const w = typeof window !== "undefined" ? window : {};
  w.OPERATION_AUDIT_FORM = OPERATION_AUDIT_FORM;
})();
