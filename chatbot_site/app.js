const chatLog = document.getElementById('chat-log');
const chatForm = document.getElementById('chat-form');
const chatInput = document.getElementById('chat-input');
const faqButtonsContainer = document.querySelector('.faq-buttons');
const serviceRadios = document.querySelectorAll('input[name="service"]');

// FAQボタンの質問と対応する画像のマップ
// 値は文字列（1枚）または配列（複数枚）で指定
const REPLY_IMAGES = {
  'フィットネスの料金を教えてください':          '/images/price-fitness.png',
  'スイミングの料金を教えてください':            ['/images/price-swimming.png', '/images/price-swimming-elite.png'],
  'フィットネスのキャンペーンを教えてください':   '/images/campaign-fitness.png',
  'スイミングのキャンペーンを教えてください':     '/images/campaign-swimming2.png',
  '短期教室について教えてください':              '/images/short-term2.png',
  'バスについて教えてください':                  ['/images/bus1.jpg', '/images/bus2.jpg'],
};

// サービスごとのFAQボタン定義
const FAQ_BUTTONS = {
  adult: [
    { emoji: '✨', label: '入会案内',        q: '入会したいです' },
    { emoji: '👀', label: '見学・体験予約',  q: '見学したいです' },
    { emoji: '💰', label: '料金について',    q: 'フィットネスの料金を教えてください' },
    { emoji: '🕐', label: '営業時間',        q: '営業時間を教えてください' },
    { emoji: '🎁', label: 'キャンペーン',    q: 'フィットネスのキャンペーンを教えてください' },
    { emoji: '🏛', label: '施設案内',        q: 'フィットネスの施設を教えてください' },
  ],
  kids: [
    { emoji: '✨', label: '入会案内',        q: '入会したいです' },
    { emoji: '👀', label: '見学・体験予約',  q: 'スイミングの見学・体験予約をしたいです' },
    { emoji: '💰', label: '料金について',    q: 'スイミングの料金を教えてください' },
    { emoji: '🕐', label: '営業時間',        q: '営業時間を教えてください' },
    { emoji: '🎁', label: 'キャンペーン',    q: 'スイミングのキャンペーンを教えてください' },
    { emoji: '📚', label: '短期教室',        q: '短期教室について教えてください' },
    { emoji: '🚌', label: 'バス案内',        q: 'バスについて教えてください' },
  ],
};

function getSelectedService() {
  const selected = document.querySelector('input[name="service"]:checked');
  return selected ? selected.value : 'adult';
}

function updateBackground(service) {
  document.body.classList.remove('service-adult', 'service-kids');
  document.body.classList.add(`service-${service}`);
}

function renderFaqButtons(service) {
  faqButtonsContainer.innerHTML = '';
  FAQ_BUTTONS[service].forEach(({ emoji, label, q }) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = `${emoji} ${label}`;
    btn.addEventListener('click', () => {
      addMessage(q, 'user');
      addMessage(getReply(q), 'bot');
      const imgs = REPLY_IMAGES[q];
      if (imgs) {
        const list = Array.isArray(imgs) ? imgs : [imgs];
        list.forEach(url => addImageMessage(url));
      }
    });
    faqButtonsContainer.appendChild(btn);
  });
}

function addMessage(text, type) {
  const div = document.createElement('div');
  div.className = `message ${type}`;
  div.innerHTML = text
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/(https?:\/\/[^\s]+|\/[\w\-.]+\.html)/g, '<a href="$1" target="_blank">$1</a>')
    .replace(/(\d{2,4}-\d{2,4}-\d{4})/g, '<a href="tel:$1">$1</a>')
    .replace(/\n/g, '<br>');
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function addImageMessage(imageUrl) {
  const div = document.createElement('div');
  div.className = 'message bot';
  const img = document.createElement('img');
  img.src = imageUrl;
  img.alt = '案内画像';
  img.className = 'reply-image';
  div.appendChild(img);
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function getReply(question) {
  const q = question;
  const service = getSelectedService();

  // キーワードが明示されている場合はそちらを優先し、なければ選択中のサービスを使う
  const mentionsAdult = /大人|フィットネス|トレーニング|ジム/.test(q);
  const mentionsKids  = /子供|キッズ|スイミング|水泳|水着/.test(q);
  const isAdult = mentionsAdult || (!mentionsKids && service === 'adult');
  const isKids  = mentionsKids  || (!mentionsAdult && service === 'kids');

  // キャンペーン
  if (/キャンペーン/.test(q)) {
    if (isAdult) return '大人のフィットネスの最新キャンペーンはこちらです。\nお電話でもご案内します：086-952-2855';
    if (isKids)  return '子供のスイミングの最新キャンペーンはこちらです。\nお電話でもご案内します：086-952-2855';
    return '最新キャンペーン情報はお電話にてご案内します。\n📞 086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
  }

  // 短期教室
  if (/短期教室|短期|期間限定/.test(q)) {
    return '短期教室についてはこちらをご確認ください。\nお電話でもご案内します：086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
  }

  // 料金
  if (/料金|費用|いくら|値段|月会費|月額/.test(q)) {
    if (isAdult) return '大人のフィットネスの料金はこちらをご確認ください。\n入会金0円・事務手数料4,400円（税込）で、月会費はプランにより異なります。';
    if (isKids)  return '子供のスイミングの料金はこちらをご確認ください。\n入会金0円で、月会費はプランにより異なります。';
    return '料金については、サービスを選択してからもう一度お聞きください 😊\nまたは直接お電話でもご案内します：086-952-2855';
  }

  // 営業時間
  if (/営業時間|何時|時間|open|閉館|開館/.test(q)) {
    if (isAdult) return '大人のフィットネスの営業時間はこちらです。\n平日: 10:00〜21:20\n土日祝: 10:00〜18:00\n※毎週木曜日は休館日です。';
    if (isKids)  return '子供のスイミングの営業時間はこちらです。\n平日: 10:00〜21:20\n土日祝: 10:00〜18:00\n※毎週木曜日は休館日です。';
    return '営業時間のご案内です。\n平日: 10:00〜21:20\n土日祝: 10:00〜18:00\n※毎週木曜日は休館日です。';
  }

  // 施設・場所
  if (/施設|設備|場所|住所|アクセス|どこ|駐車場/.test(q)) {
    if (isAdult) return '大人のフィットネスの施設・設備はこちらをご覧ください。\nhttps://www.acro.co.jp/fitness/\n\n住所: 岡山市東区瀬戸町沖333\n駐車場: 50台（無料）\nお電話: 086-952-2855';
    return '施設案内はこちらです。\nhttps://www.acro.co.jp/\n\n住所: 岡山市東区瀬戸町沖333\n駐車場: 50台（無料）\nお電話: 086-952-2855';
  }

  // バス案内
  if (/バス/.test(q)) {
    if (isAdult) return '大人のフィットネスにはバス送迎サービスはございません。\nお電話: 086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
    return 'バスのご案内はこちらです。\nお電話でも詳しくご案内します：086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
  }

  // スイミング見学（専用リンクあり）
  if (q === 'スイミングの見学・体験予約をしたいです') {
    return '体験予約はこちらのフォームからお申し込みいただけます 😊\nhttps://www3.clubnet.ne.jp/acroport/mypage/index.php/trial\n\nお電話でも受け付けております。\nお電話: 086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
  }

  // 見学（フィットネス・共通）
  if (/見学|体験|見てみ/.test(q)) {
    return '見学・無料体験は随時受付中です！\nお電話またはフロントにてお申し込みください。\nお電話: 086-952-2855\n（平日10:00〜21:20 / 土日祝10:00〜18:00）';
  }

  // 入会
  if (/入会|加入|申し込み|登録/.test(q)) {
    return '入会手続きはこちらからお願いします。\nhttps://www3.clubnet.ne.jp/acroport/entry/\nお気軽にどうぞ！\n\nお電話: 086-952-2855\n（当日入会も可能です）';
  }

  // 電話・問い合わせ
  if (/電話|問い合わせ|連絡|tel/.test(q)) {
    return 'お電話でのお問い合わせはこちらです。\n📞 086-952-2855\n平日: 10:00〜21:20\n土日祝: 10:00〜18:00\n※木曜日は休館日です。';
  }

  // 大人フィットネス全般
  if (isAdult) {
    return '大人のフィットネスについて詳しくはこちらをご覧ください。\nhttps://www.acro.co.jp/fitness/\n\nご不明な点はお電話でもどうぞ：086-952-2855';
  }

  // 子供スイミング全般
  if (isKids) {
    return '子供のスイミングについて詳しくはこちらをご覧ください。\nhttps://www.acro.co.jp/swimming/\n\nご不明な点はお電話でもどうぞ：086-952-2855';
  }

  return 'ご質問ありがとうございます 😊\nサービスを選択いただくか、もう少し詳しく教えていただけますか？\nお電話でもお気軽にどうぞ：086-952-2855';
}

chatForm.addEventListener('submit', (e) => {
  e.preventDefault();
  const q = chatInput.value.trim();
  if (!q) return;
  chatInput.value = '';
  addMessage(q, 'user');
  addMessage(getReply(q), 'bot');
});

serviceRadios.forEach((radio) => {
  radio.addEventListener('change', () => {
    const s = getSelectedService();
    const label = s === 'adult' ? '大人のフィットネス' : '子供のスイミング';
    addMessage(`サービスを「${label}」に設定しました。`, 'bot');
    renderFaqButtons(s);
    updateBackground(s);
  });
});

// 初期表示（デフォルトは大人のフィットネス）
updateBackground('adult');
renderFaqButtons('adult');
addMessage('こんにちは！アクロポートガイドです 😊\nサービスを選んでからご質問ください！', 'bot');
