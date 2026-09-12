/**
 * BOT WHATSAPP - QUIZ DE CULTURE GENERALE
 * ----------------------------------------
 * Ce fichier fait tourner le bot. Tu n'as normalement pas besoin d'y toucher.
 * Toutes les questions viennent du fichier questions.xlsx (dans ce même dossier).
 */
 
const {
  default: makeWASocket,
  useMultiFileAuthState,
  fetchLatestBaileysVersion,
  DisconnectReason,
} = require('@whiskeysockets/baileys');
const P = require('pino');
const fs = require('fs');
const path = require('path');
const http = require('http');
const XLSX = require('xlsx');
const sharp = require('sharp');
const readline = require('readline');
 
// ---------- SERVEUR HTTP (obligatoire pour Render Web Service) ----------
// Render exige qu'un service "Web Service" réponde sur le port qu'il fournit,
// sinon il considère le déploiement en échec et arrête le service.
// Ce serveur ne fait rien d'autre que répondre "OK" — le bot WhatsApp tourne à côté.
const PORT = process.env.PORT || 3000;
http
  .createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('OK - Bot Quiz WhatsApp actif');
  })
  .listen(PORT, () => {
    console.log(`🌐 Serveur HTTP démarré sur le port ${PORT} (health check Render)`);
  });
 
const QUESTIONS_FILE = path.join(__dirname, 'questions.xlsx');
const SCORES_FILE = path.join(__dirname, 'scores.json');
const AUTH_FOLDER = path.join(__dirname, 'auth_info');
 
// Ton numero WhatsApp (celui que le bot va utiliser), format international SANS le +
// Exemple pour la Cote d'Ivoire : "2250700000000"
const PHONE_NUMBER = process.env.BOT_PHONE_NUMBER || '';
 
// Numero(s) autorisés à lancer/arrêter le quiz et créer des stickers.
// Par défaut : uniquement ton propre numéro (celui du bot). Pour autoriser
// d'autres personnes, mets par exemple : OWNER_NUMBERS=2250700000000,2250100000000
const OWNER_NUMBERS = (process.env.OWNER_NUMBERS || PHONE_NUMBER)
  .split(',')
  .map((n) => n.trim())
  .filter(Boolean);
 
function extraireNumero(id) {
  // Un id WhatsApp ressemble à "2250700000000@s.whatsapp.net" ou "2250700000000:12@s.whatsapp.net"
  return String(id).split('@')[0].split(':')[0];
}
 
function estProprietaire(senderId) {
  const numero = extraireNumero(senderId);
  const autorise = OWNER_NUMBERS.includes(numero);
  if (!autorise) {
    console.log(
      `🔍 DEBUG owner check → senderId reçu: "${senderId}" | numéro extrait: "${numero}" | OWNER_NUMBERS configurés: [${OWNER_NUMBERS.join(', ')}]`
    );
  }
  return autorise;
}
 
// ---------- GESTION DES SCORES ----------
function loadScores() {
  if (!fs.existsSync(SCORES_FILE)) return {};
  try {
    return JSON.parse(fs.readFileSync(SCORES_FILE, 'utf-8'));
  } catch {
    return {};
  }
}
function saveScores(scores) {
  fs.writeFileSync(SCORES_FILE, JSON.stringify(scores, null, 2));
}
 
// ---------- CHARGEMENT DES QUESTIONS ----------
function loadQuestions() {
  if (!fs.existsSync(QUESTIONS_FILE)) {
    console.log('⚠️  Fichier questions.xlsx introuvable. Ajoute-le dans ce dossier.');
    return [];
  }
  const workbook = XLSX.readFile(QUESTIONS_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
 
  return rows
    .filter((r) => String(r.Question || '').trim() !== '')
    .map((r) => ({
      categorie: String(r.Categorie || 'Général').trim(),
      question: String(r.Question).trim(),
      bonneReponse: String(r.BonneReponse || '').trim(),
    }));
}
 
// Compare deux textes en ignorant majuscules/minuscules, accents et ponctuation
function normaliser(texte) {
  return String(texte)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // enlève les accents
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, '') // enlève la ponctuation
    .trim()
    .replace(/\s+/g, ' ');
}
 
let QUESTIONS = loadQuestions();
console.log(`✅ ${QUESTIONS.length} questions chargées.`);
 
// Etat des quiz en cours, par discussion (groupe ou privé)
// { [chatId]: { question, dejaRepondu: Set(userId), resolved, timeoutId } }
const activeQuizzes = {};
 
// Etat de la "session" de quiz enchaîné par discussion : { [chatId]: { running: true/false } }
const sessions = {};
 
const TEMPS_REPONSE_MS = 12000; // 12 secondes pour répondre
const DELAI_PROCHAINE_QUESTION_MS = 4000; // 4 secondes de pause avant la question suivante
 
function pickRandomQuestion() {
  if (QUESTIONS.length === 0) return null;
  return QUESTIONS[Math.floor(Math.random() * QUESTIONS.length)];
}
 
// Mélange un tableau (pour éviter de reposer 2 fois la même question dans une série)
function melanger(tableau) {
  const copie = [...tableau];
  for (let i = copie.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [copie[i], copie[j]] = [copie[j], copie[i]];
  }
  return copie;
}
 
function formatQuestion(q, numero, total) {
  const compteur = total ? `(Question ${numero}/${total}) ` : '';
  return `🧠 *QUIZ CULTURE GENERALE* ${compteur}(${q.categorie})\n\n${q.question}\n\n👉 Tape directement ta réponse dans le chat. Tu as 12 secondes !`;
}
 
function formatScore(chatId) {
  const scores = loadScores();
  const chatScores = scores[chatId] || {};
  const entries = Object.entries(chatScores).sort((a, b) => b[1].points - a[1].points);
  if (entries.length === 0) return 'Aucun score enregistré pour le moment.';
  let text = '🏆 *Classement*\n\n';
  entries.forEach(([, data], i) => {
    text += `${i + 1}. ${data.nom} — ${data.points} pt(s)\n`;
  });
  return text;
}
 
// ---------- DEMARRAGE DU BOT ----------
async function startBot() {
  const { state, saveCreds } = await useMultiFileAuthState(AUTH_FOLDER);
  const { version } = await fetchLatestBaileysVersion();
 
  const sock = makeWASocket({
    version,
    auth: state,
    logger: P({ level: 'silent' }),
    printQRInTerminal: false,
  });
 
  // ---- Connexion par CODE (pas besoin de scanner un QR) ----
  if (!sock.authState.creds.registered) {
    if (!PHONE_NUMBER) {
      console.log('❌ Ajoute ton numéro dans la variable BOT_PHONE_NUMBER (voir README).');
      process.exit(1);
    }
    setTimeout(async () => {
      try {
        const code = await sock.requestPairingCode(PHONE_NUMBER);
        console.log('\n=========================================');
        console.log(` TON CODE DE CONNEXION WHATSAPP : ${code}`);
        console.log(' Ouvre WhatsApp > Paramètres > Appareils connectés');
        console.log(' > Connecter un appareil > "Se connecter avec un numéro"');
        console.log(' et entre ce code.');
        console.log('=========================================\n');
      } catch (e) {
        console.log('Erreur lors de la demande du code :', e);
      }
    }, 3000);
  }
 
  sock.ev.on('creds.update', saveCreds);
 
  // ---- Pose une question et lance le minuteur de 12 secondes ----
  async function poserQuestion(chatId) {
    const session = sessions[chatId];
    if (!session || !session.running) return;
 
    // Si un nombre de questions était fixé (!quiz30) et qu'on l'a atteint : fin automatique
    if (session.total !== null && session.count >= session.total) {
      session.running = false;
      await sock.sendMessage(chatId, {
        text: `🏁 Quiz terminé ! ${session.total} question(s) posée(s).`,
      });
      await sock.sendMessage(chatId, { text: formatScore(chatId) });
      return;
    }
 
    // On pioche dans la réserve mélangée pour ne pas répéter une question deux fois
    let q;
    if (session.pool && session.pool.length > 0) {
      q = session.pool.pop();
    } else {
      q = pickRandomQuestion();
    }
 
    if (!q) {
      await sock.sendMessage(chatId, {
        text: '⚠️ Aucune question disponible. Vérifie le fichier questions.xlsx.',
      });
      session.running = false;
      await sock.sendMessage(chatId, { text: formatScore(chatId) });
      return;
    }
 
    session.count += 1;
    activeQuizzes[chatId] = { question: q, resolved: false };
    await sock.sendMessage(chatId, { text: formatQuestion(q, session.count, session.total) });
 
    const timeoutId = setTimeout(async () => {
      const quiz = activeQuizzes[chatId];
      if (quiz && !quiz.resolved) {
        quiz.resolved = true;
        await sock.sendMessage(chatId, {
          text: `⏰ Temps écoulé ! La bonne réponse était : *${quiz.question.bonneReponse}*`,
        });
        delete activeQuizzes[chatId];
        if (sessions[chatId]?.running) {
          setTimeout(() => poserQuestion(chatId), DELAI_PROCHAINE_QUESTION_MS);
        }
      }
    }, TEMPS_REPONSE_MS);
 
    sessions[chatId].timeoutId = timeoutId;
  }
 
  sock.ev.on('connection.update', (update) => {
    const { connection, lastDisconnect } = update;
    if (connection === 'close') {
      const shouldReconnect =
        lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
      console.log('Connexion fermée. Reconnexion :', shouldReconnect);
      if (shouldReconnect) startBot();
    } else if (connection === 'open') {
      console.log('✅ Bot connecté à WhatsApp !');
    }
  });
 
  // ---- Reception des messages ----
  sock.ev.on('messages.upsert', async ({ messages, type }) => {
    if (type !== 'notify') return;
    for (const msg of messages) {
      await traiterMessage(msg);
    }
  });
 
  async function traiterMessage(msg) {
    if (!msg.message) return;
 
    const chatId = msg.key.remoteJid;
    const senderId = msg.key.participant || msg.key.remoteJid;
    const senderName = msg.pushName || 'Joueur';
 
    const textMessage =
      msg.message.conversation ||
      msg.message.extendedTextMessage?.text ||
      msg.message.imageMessage?.caption ||
      '';
    const bodyLower = textMessage.trim().toLowerCase();
 
    try {
      // ---------- COMMANDE : !aide ----------
      if (bodyLower === '!aide' || bodyLower === '!help') {
        await sock.sendMessage(chatId, {
          text:
            '🤖 *Commandes disponibles*\n\n' +
            '!quiz → démarrer une série de 25 questions, avec classement final automatique\n' +
            '!quiz30 → démarrer une série de 30 questions précises (ou tout autre nombre)\n' +
            '!stop → arrêter la série de questions et afficher le score final\n' +
            '!score → voir le classement de la discussion\n' +
            '!sticker → réponds à une image avec ce mot pour la transformer en sticker (réservé à l’organisateur)\n' +
            '!aide → afficher ce message',
        });
        return;
      }
 
      // ---------- COMMANDE : !quiz ou !quiz30 (nombre de questions) ----------
      // Ouvert à tout le monde (plus de restriction organisateur)
      const matchQuiz = bodyLower.match(/^!quiz(\d+)?$/);
      if (matchQuiz) {
        if (sessions[chatId]?.running) {
          await sock.sendMessage(chatId, { text: 'Un quiz est déjà en cours ! Tape !stop pour l’arrêter.' });
          return;
        }
        const total = matchQuiz[1] ? parseInt(matchQuiz[1], 10) : 25;
        if (total !== null && total <= 0) {
          await sock.sendMessage(chatId, { text: 'Indique un nombre de questions valide, ex : !quiz30' });
          return;
        }
        sessions[chatId] = {
          running: true,
          timeoutId: null,
          total,
          count: 0,
          pool: melanger(QUESTIONS),
        };
        await poserQuestion(chatId);
        return;
      }
 
      // ---------- COMMANDE : !stop ----------
      // Ouvert à tout le monde (plus de restriction organisateur)
      if (bodyLower === '!stop') {
        if (sessions[chatId]) {
          sessions[chatId].running = false;
          if (sessions[chatId].timeoutId) clearTimeout(sessions[chatId].timeoutId);
        }
        delete activeQuizzes[chatId];
        await sock.sendMessage(chatId, { text: '🛑 Quiz arrêté.' });
        await sock.sendMessage(chatId, { text: formatScore(chatId) });
        return;
      }
 
      // ---------- COMMANDE : !score ----------
      if (bodyLower === '!score') {
        await sock.sendMessage(chatId, { text: formatScore(chatId) });
        return;
      }
 
      // ---------- COMMANDE : !sticker (réponse à une image) ----------
      const isReplyToImage =
        msg.message.imageMessage ||
        msg.message.extendedTextMessage?.contextInfo?.quotedMessage?.imageMessage;
 
      if (bodyLower === '!sticker' && isReplyToImage) {
        if (!estProprietaire(senderId)) {
          await sock.sendMessage(chatId, { text: '🔒 Seul l’organisateur peut créer des stickers.' });
          return;
        }
        await sock.sendMessage(chatId, { text: '🎨 Création du sticker...' });
 
        let imageMessageContent = msg.message.imageMessage;
        let messageForDownload = msg;
 
        if (!imageMessageContent) {
          // L'image est dans le message cité (reply)
          const quoted = msg.message.extendedTextMessage.contextInfo;
          imageMessageContent = quoted.quotedMessage.imageMessage;
          messageForDownload = {
            key: {
              remoteJid: chatId,
              id: quoted.stanzaId,
              fromMe: false,
              participant: quoted.participant,
            },
            message: quoted.quotedMessage,
          };
        }
 
        const { downloadMediaMessage } = require('@whiskeysockets/baileys');
        const buffer = await downloadMediaMessage(messageForDownload, 'buffer', {});
 
        const webpBuffer = await sharp(buffer)
          .resize(512, 512, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
          .webp()
          .toBuffer();
 
        await sock.sendMessage(chatId, { sticker: webpBuffer });
        return;
      }
 
      // ---------- REPONSE A UNE QUESTION EN COURS ----------
      const quiz = activeQuizzes[chatId];
      const estUneCommande = bodyLower.startsWith('!');
 
      // Note : on ne bloque plus un joueur après une mauvaise réponse — il peut
      // réessayer autant de fois qu'il veut pendant les 12 secondes. Seul le
      // drapeau "resolved" (une bonne réponse déjà trouvée) arrête le jeu.
      if (quiz && !quiz.resolved && !msg.key.fromMe && textMessage.trim() !== '' && !estUneCommande) {
        const bonne = normaliser(quiz.question.bonneReponse);
        const proposee = normaliser(textMessage);
        const motsDeLaBonneReponse = bonne.split(' ').filter((m) => m.length > 1);
        const estCorrecte =
          proposee === bonne ||
          motsDeLaBonneReponse.includes(proposee) || // ex: "Einstein" au lieu de "Albert Einstein"
          (bonne.length > 3 && proposee.includes(bonne));
 
        if (estCorrecte) {
          quiz.resolved = true;
          if (sessions[chatId]?.timeoutId) clearTimeout(sessions[chatId].timeoutId);
 
          const scores = loadScores();
          if (!scores[chatId]) scores[chatId] = {};
          if (!scores[chatId][senderId]) scores[chatId][senderId] = { nom: senderName, points: 0 };
          scores[chatId][senderId].points += 10;
          scores[chatId][senderId].nom = senderName;
          saveScores(scores);
 
          await sock.sendMessage(chatId, {
            text: `✅ Bonne réponse, ${senderName} ! (+10 points)`,
          });
          delete activeQuizzes[chatId];
 
          if (sessions[chatId]?.running) {
            setTimeout(() => poserQuestion(chatId), DELAI_PROCHAINE_QUESTION_MS);
          }
        } else {
          await sock.sendMessage(chatId, {
            text: `❌ Mauvaise réponse, ${senderName}. Essaie encore !`,
          });
        }
      }
    } catch (err) {
      console.log('Erreur en traitant un message :', err);
    }
  }
}
 
startBot();
