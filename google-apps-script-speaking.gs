const SHEET_NAME = 'Speaking Level Check';
const KIE_CHAT_COMPLETIONS_URL = 'https://api.kie.ai/gemini-3-5-flash-openai/v1/chat/completions';
const KIE_FILE_UPLOAD_URL = 'https://kieai.redpandaai.co/api/file-base64-upload';
const SPEAKING_AUDIO_RETENTION_DAYS = 1;
const SPEAKING_ARTIFACT_RETENTION_DAYS = 30;
const SPEAKING_ARTIFACT_FOLDER_NAME = 'Speaking Level Check Artifacts';
const HEADERS = [
  'timestamp', 'name', 'whatsapp', 'social_media_account', 'followers_range', 'age', 'status',
  'score', 'level', 'task_response', 'fluency_coherence', 'pronunciation',
  'grammar', 'vocabulary', 'communication_strategy', 'zero_tasks', 'teacher_note',
  'certificate_url', 'full_report_pdf_url', 'artifact_expires_at'
];

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  if (payload.action === 'scoreSpeaking') {
    const speakingScore = scoreSpeakingWithGemini(payload);
    return jsonResponse({ ok: !speakingScore.error, speakingScore });
  }
  if (payload.action === 'saveSpeakingResult') {
    return jsonResponse({ ok: true, resultSave: saveSpeakingResult(payload) });
  }
  if (payload.action === 'saveSpeakingArtifacts') {
    return jsonResponse({ ok: true, artifacts: saveSpeakingArtifacts(payload) });
  }
  return jsonResponse({ ok: false, error: 'Unknown action' });
}

function authorizeSpeakingBackend() {
  const sheet = getSheet();
  const testFile = DriveApp.createFile(
    Utilities.newBlob('authorization check', 'text/plain', `speaking-auth-check-${Date.now()}.txt`)
  );
  testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  testFile.setTrashed(true);

  const response = UrlFetchApp.fetch('https://api.kie.ai', { muteHttpExceptions: true });
  return {
    ok: true,
    sheet: sheet.getName(),
    drive: 'authorized',
    urlFetchStatus: response.getResponseCode()
  };
}

function scoreSpeakingWithGemini(payload) {
  const keys = getKieApiKeys();
  if (!keys.length) return { error: 'KIE_API_KEYS belum diatur di Script Properties.', source: 'kie configuration error' };
  cleanupOldSpeakingFiles();

  const tasks = (payload.tasks || []).map((task) => ({
    id: String(task.id || ''),
    title: String(task.title || ''),
    prompt: String(task.prompt || ''),
    imageContext: String(task.imageContext || ''),
    targetSeconds: Number(task.targetSeconds || 0),
    duration: Number(task.duration || 0),
    mimeType: String(task.mimeType || 'audio/webm'),
    audioBase64: String(task.audioBase64 || '')
  }));

  if (!tasks.some((task) => task.audioBase64)) return zeroResult('no audio submitted');

  try {
    const keyState = { keys, blocked: {} };
    const taskScores = tasks.map((task) => scoreSingleTaskWithGemini(keyState, task));
    const result = buildSpeakingResultFromTasks(taskScores);
    result.source = 'kie.ai gemini-3-5-flash-openai';
    return result;
  } catch (error) {
    return {
      source: 'kie api error',
      error: String(error.message || error),
      debug: {
        message: String(error.message || error),
        body: String(error.body || '').slice(0, 1000),
        audioFormat: error.audioFormat || '',
        mimeType: error.mimeType || ''
      }
    };
  }
}

function getKieApiKeys() {
  return getScriptPropertyKeys('KIE_API_KEYS', 'KIE_API_KEY');
}

function getScriptPropertyKeys(multiName, singleName) {
  const props = PropertiesService.getScriptProperties();
  const multi = props.getProperty(multiName) || '';
  const single = props.getProperty(singleName) || '';
  return `${multi}\n${single}`
    .split(/[\n,;]/)
    .map((key) => key.trim())
    .filter((key, index, list) => key && list.indexOf(key) === index);
}

function scoreSingleTaskWithGemini(keyState, task) {
  if (!task.audioBase64 || task.duration < 2) {
    return zeroTaskScore(task, 'Tidak ada audio valid untuk task ini.');
  }

  const audioActivity = analyzeWavActivity(task);
  if (audioActivity.silent) {
    return zeroTaskScore(task, 'Rekaman hening, jadi tidak ada speaking yang bisa dinilai.');
  }
  if (!audioActivity.hasSpeechLikeSignal) {
    throw new Error('Rekaman terlalu pelan atau tidak cukup jelas untuk dinilai. Silakan rekam ulang tes.');
  }

  const prompt = [
    'You are an expert English speaking placement examiner for an Indonesian speaking-focused English course.',
    'Assess this ONE submitted speaking recording using CEFR-aligned placement judgement.',
    'Important zero-score rule: if this task has silent audio, mostly noise, non-speech, or the learner says almost nothing meaningful, score must be 0 and must not be guessed from the prompt.',
    'Do not reward reading the prompt silently. Score only audible spoken English.',
    'For picture description tasks, use imageContext as the reference for what the learner was looking at. Do not assume a different picture.',
    'Use a stricter placement rubric. Do not be generous just because the answer is understandable.',
    'Rubric weights for overall scoring are: content/relevance via taskResponse 10%, fluencyCoherence 25%, pronunciation 25%, grammar 20%, vocabulary 20%. communicationStrategy is a legacy field and must not inflate the overall score.',
    'Answer length is evidence, not a score by itself. However, very short answers cannot receive high scores because there is not enough evidence of speaking ability.',
    'For productive tasks (self introduction, daily life, picture description, opinion), content/relevance is a gatekeeper. Pronunciation or fluency must not rescue a short, vague, off-topic, or nonsense answer.',
    'Strict cap rules for productive tasks: silence/no meaningful English = 0; 1-3 English words = max 20/100; one very short sentence = max 35/100; 1-2 short/simple sentences with little detail = max 45/100; two simple but relevant sentences = max 50/100; vague/general answer with no specific details = max 55/100; mostly understandable but very basic grammar/vocabulary = max 60/100; repeated basic grammar errors = max 60/100; many fillers, long pauses, or broken flow = max 65/100; incomplete or partly off-topic answer = max 45/100; clearly off-topic/ngawur answer = max 30/100.',
    'If taskResponse is below 40%, displayed score must be max 45 no matter how clear the pronunciation is. If taskResponse is below 30%, displayed score must be max 35. If taskResponse is below 20%, displayed score must be max 25.',
    'If the answer is only 1-2 sentences and does not directly answer the prompt, score it around 20-35, not 50+.',
    '70+ is only for answers that are clear, sufficiently developed, relevant, reasonably fluent, and fairly stable in grammar and pronunciation. 80+ requires natural delivery, enough detail, varied vocabulary, and only minor errors. 90+ requires very natural, accurate, well-structured speaking with minimal issues.',
    'Advanced vocabulary is not required for a good score. Precise, natural, and varied vocabulary matters more. Advanced words used incorrectly should reduce vocabulary and/or grammar.',
    'For read-aloud and repeat-sentence tasks, content/relevance can be high if the expected text is attempted, but the raw score must still depend mainly on pronunciation, fluency, rhythm, intonation, and accuracy.',
    'Give score as a raw 0-20 task score that matches the strict 0-100 judgement divided by 5. Example: 60/100 -> score 12, 70/100 -> score 14, 85/100 -> score 17.',
    'Also give criteriaPercent values from 0-100 for taskResponse, fluencyCoherence, pronunciation, grammar, vocabulary, and communicationStrategy. Set communicationStrategy to 0 unless you have a specific strategy observation.',
    'If the displayed score is below 100, every feedback must explain the remaining gap clearly. Do not write "no meaningful mistake" as the only correction for a score below 100. Mention the specific reason: stress, intonation, rhythm, connected speech, pace control, consistency, minor pronunciation clarity, or confidence.',
    'For scores 85-95, the feedback should say that the answer is strong, but still explain what kept it from 100 in a specific and understandable way.',
    'Write all human-facing feedback in friendly, casual Bahasa Indonesia using "kamu", not "Anda". Keep transcript exactly as spoken.',
    'Tone: Gen-Z friendly, fun, supportive, and clear, like a helpful speaking coach. Use phrases such as "ini udah oke", "next kamu bisa...", "biar makin natural", "bagian ini perlu dipoles", but do not overdo slang and do not use emojis.',
    'Give constructive, answer-specific feedback. Do not use generic template feedback. Mention what was good from the learner answer, what was wrong or unclear, how to correct grammar, how to correct pronunciation if you hear a mispronunciation, a better/natural sentence example, and why the task got that score. In scoreReason, explain the displayed 0-100 score, not only the raw 0-20 score.',
    'For grammarFix and pronunciationFix, write short separate correction sentences, not one long paragraph. Example: "Seharusnya ... (bukan ...)." "Next, pakai ... biar lebih natural."',
    'If pronunciation cannot be judged confidently, set pronunciationFix to "Tidak ada koreksi pronunciation spesifik yang terdengar jelas." Do not invent word-level pronunciation errors.',
    'Use encouraging but honest wording. Avoid stiff examiner language, avoid "peserta", "pembelajar", and avoid formal words like "Anda".',
    'Give strict JSON only with this shape:',
    '{"id":"","score":0,"transcript":"","note":"","good":"","correction":"","grammarFix":"","pronunciationFix":"","betterAnswer":"","scoreReason":"","isZero":true,"criteriaPercent":{"taskResponse":0,"fluencyCoherence":0,"pronunciation":0,"grammar":0,"vocabulary":0,"communicationStrategy":0}}',
    'Task metadata:',
    JSON.stringify({
      id: task.id,
      title: task.title,
      prompt: task.prompt,
      imageContext: task.imageContext,
      targetSeconds: task.targetSeconds,
      duration: task.duration,
      mimeType: task.mimeType
    })
  ].join('\n');

  return callKieWithBackupKeys(keyState, prompt, task);
}

function analyzeWavActivity(task) {
  if (!String(task.mimeType || '').toLowerCase().includes('wav')) {
    throw new Error('Format audio belum dapat diverifikasi. Silakan rekam ulang tes.');
  }
  const bytes = Utilities.base64Decode(task.audioBase64);
  const byte = (index) => bytes[index] & 255;
  const label = (index) => String.fromCharCode(byte(index), byte(index + 1), byte(index + 2), byte(index + 3));
  const uint16 = (index) => byte(index) | (byte(index + 1) << 8);
  const uint32 = (index) => (byte(index) | (byte(index + 1) << 8) | (byte(index + 2) << 16) | (byte(index + 3) << 24)) >>> 0;
  if (bytes.length < 44 || label(0) !== 'RIFF' || label(8) !== 'WAVE') {
    throw new Error('Rekaman WAV tidak valid. Silakan rekam ulang tes.');
  }

  let sampleRate = 0;
  let dataStart = -1;
  let dataLength = 0;
  for (let offset = 12; offset + 8 <= bytes.length;) {
    const size = uint32(offset + 4);
    const start = offset + 8;
    if (size > bytes.length - start) throw new Error('Rekaman WAV terpotong. Silakan rekam ulang tes.');
    if (label(offset) === 'fmt ') {
      if (size < 16 || uint16(start) !== 1 || uint16(start + 2) !== 1 || uint16(start + 14) !== 16) {
        throw new Error('Rekaman harus berupa WAV PCM mono 16-bit. Silakan rekam ulang tes.');
      }
      sampleRate = uint32(start + 4);
    } else if (label(offset) === 'data') {
      dataStart = start;
      dataLength = size;
    }
    offset = start + size + (size & 1);
  }
  if (!sampleRate || dataStart < 0 || dataLength < sampleRate * 2) {
    throw new Error('Rekaman tidak lengkap. Silakan rekam ulang tes.');
  }

  const frameSamples = Math.max(1, Math.round(sampleRate * 0.02));
  const frameLevels = [];
  let peak = 0;
  for (let offset = dataStart; offset + 1 < dataStart + dataLength;) {
    let sumSquares = 0;
    let count = 0;
    for (; count < frameSamples && offset + 1 < dataStart + dataLength; count += 1, offset += 2) {
      const unsigned = uint16(offset);
      const sample = unsigned >= 32768 ? unsigned - 65536 : unsigned;
      const level = Math.abs(sample) / 32768;
      peak = Math.max(peak, level);
      sumSquares += level * level;
    }
    if (count) frameLevels.push(Math.sqrt(sumSquares / count));
  }
  const activeFrames = frameLevels.filter((level) => level >= 0.008).length;
  const sortedLevels = frameLevels.slice().sort((a, b) => a - b);
  const low = sortedLevels[Math.floor(sortedLevels.length * 0.1)] || 0;
  const high = sortedLevels[Math.floor(sortedLevels.length * 0.9)] || 0;
  return {
    silent: peak < 0.002 && activeFrames < 5,
    hasSpeechLikeSignal: peak >= 0.015 && activeFrames >= 10 && high >= 0.012 && high >= Math.max(low * 1.5, low + 0.005)
  };
}

function callKieWithBackupKeys(keyState, prompt, task) {
  let lastError = null;
  const availableKeys = keyState.keys.filter((key) => !keyState.blocked[key]);
  if (!availableKeys.length) {
    throw new Error('All Kie backup keys are blocked by quota/rate limit for this test.');
  }
  for (let index = 0; index < availableKeys.length; index += 1) {
    const key = availableKeys[index];
    try {
      return callKieOnce(key, prompt, task);
    } catch (error) {
      lastError = error;
      if (!isRetryableKieError(error)) {
        throw error;
      }
      keyState.blocked[key] = true;
      if (index === availableKeys.length - 1) {
        throw error;
      }
      Utilities.sleep(350);
    }
  }
  throw lastError || new Error('Kie API failed with all backup keys.');
}

function callKieOnce(key, prompt, task) {
  const audioFormat = getAudioFormat(task.mimeType);
  try {
    const audioUrl = uploadAudioToKie(key, task);
    const response = UrlFetchApp.fetch(KIE_CHAT_COMPLETIONS_URL, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${key}` },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        messages: [{
          role: 'user',
          content: [
            { type: 'text', text: prompt },
            { type: 'image_url', image_url: { url: audioUrl } }
          ]
        }],
        temperature: 0.05,
        response_format: {
          type: 'json_schema',
          json_schema: {
            name: 'speaking_task_score',
            strict: true,
            schema: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                score: { type: 'number' },
                transcript: { type: 'string' },
                note: { type: 'string' },
                good: { type: 'string' },
                correction: { type: 'string' },
                grammarFix: { type: 'string' },
                pronunciationFix: { type: 'string' },
                betterAnswer: { type: 'string' },
                scoreReason: { type: 'string' },
                isZero: { type: 'boolean' },
                criteriaPercent: {
                  type: 'object',
                  properties: {
                    taskResponse: { type: 'number' },
                    fluencyCoherence: { type: 'number' },
                    pronunciation: { type: 'number' },
                    grammar: { type: 'number' },
                    vocabulary: { type: 'number' },
                    communicationStrategy: { type: 'number' }
                  },
                  required: ['taskResponse', 'fluencyCoherence', 'pronunciation', 'grammar', 'vocabulary', 'communicationStrategy']
                }
              },
              required: ['id', 'score', 'transcript', 'note', 'good', 'correction', 'grammarFix', 'pronunciationFix', 'betterAnswer', 'scoreReason', 'isZero', 'criteriaPercent']
            }
          }
        }
      })
    });
    const responseText = response.getContentText();
    const body = JSON.parse(responseText);
    if (response.getResponseCode() >= 400 || (body.code && Number(body.code) >= 400)) {
      const error = new Error(`Kie API error ${body.code || response.getResponseCode()}: ${body.msg || responseText.slice(0, 500)}`);
      error.status = Number(body.code || response.getResponseCode());
      error.body = response.getContentText();
      error.audioFormat = audioFormat;
      error.mimeType = task.mimeType;
      throw error;
    }
    const text = body.choices?.[0]?.message?.content || '{}';
    if (!body.choices?.length || text === '{}') {
      const error = new Error(`Kie API returned no scoring candidate: ${response.getContentText().slice(0, 500)}`);
      error.status = 502;
      error.body = response.getContentText();
      error.audioFormat = audioFormat;
      error.mimeType = task.mimeType;
      throw error;
    }
    return normalizeTaskScore(task, parseModelJson(text));
  } catch (error) {
    error.audioFormat = error.audioFormat || audioFormat;
    error.mimeType = error.mimeType || task.mimeType;
    throw error;
  }
}

function getAudioFormat(mimeType) {
  const type = String(mimeType || '').toLowerCase();
  if (type.includes('wav')) return 'wav';
  if (type.includes('mp3') || type.includes('mpeg')) return 'mp3';
  if (type.includes('mp4') || type.includes('m4a') || type.includes('aac')) return 'mp4';
  if (type.includes('ogg')) return 'ogg';
  if (type.includes('webm')) return 'webm';
  return 'wav';
}

function uploadAudioToKie(key, task) {
  const mimeType = task.mimeType || 'audio/wav';
  const extension = getAudioFormat(mimeType);
  const safeId = String(task.id || 'task').replace(/[^a-z0-9_-]/gi, '-');
  const response = UrlFetchApp.fetch(KIE_FILE_UPLOAD_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${key}` },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      base64Data: `data:${mimeType};base64,${task.audioBase64}`,
      uploadPath: 'speaking-tests',
      fileName: `speaking-${safeId}-${Date.now()}.${extension}`
    })
  });
  const bodyText = response.getContentText();
  const body = JSON.parse(bodyText);
  const url = body.data?.downloadUrl || body.data?.fileUrl;
  if (response.getResponseCode() >= 400 || body.success !== true || !url) {
    const error = new Error(`Kie audio upload error ${body.code || response.getResponseCode()}: ${body.msg || bodyText.slice(0, 500)}`);
    error.status = Number(body.code || response.getResponseCode());
    error.body = bodyText;
    throw error;
  }
  return url;
}

function cleanupOldSpeakingAudioFiles() {
  try {
    const cutoff = Date.now() - audioRetentionMs();
    const files = DriveApp.searchFiles("title contains 'speaking-' and trashed = false");
    while (files.hasNext()) {
      const file = files.next();
      if (file.getDateCreated().getTime() < cutoff) file.setTrashed(true);
    }
  } catch (error) {
    console.warn(`Failed to cleanup old speaking audio files: ${error}`);
  }
}

function saveSpeakingArtifacts(payload) {
  cleanupOldSpeakingFiles();
  const user = payload.user || {};
  const artifacts = payload.artifacts || {};
  const folder = getSpeakingArtifactFolder();
  const safeName = sanitizeFilename(user.name || 'peserta');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const saved = {};

  if (artifacts.certificatePng) {
    saved.certificateUrl = saveDataUrlFile(
      folder,
      artifacts.certificatePng,
      `speaking-certificate-${safeName}-${stamp}.png`
    );
  }

  if (artifacts.fullReportPdf) {
    saved.fullReportPdfUrl = saveDataUrlFile(
      folder,
      artifacts.fullReportPdf,
      `speaking-full-report-${safeName}-${stamp}.pdf`
    );
  }

  saved.expiresAt = new Date(Date.now() + artifactRetentionMs()).toISOString();
  if (payload.rowNumber) updateSpeakingArtifactLinks(payload.rowNumber, saved);
  return saved;
}

function updateSpeakingArtifactLinks(rowNumber, artifacts) {
  const row = Number(rowNumber);
  if (!Number.isFinite(row) || row < 2) return false;
  const sheet = getSheet();
  const headers = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0].map(String);
  const updates = {
    certificate_url: artifacts.certificateUrl,
    full_report_pdf_url: artifacts.fullReportPdfUrl,
    artifact_expires_at: artifacts.expiresAt
  };

  Object.keys(updates).forEach(header => {
    if (!updates[header]) return;
    const index = headers.indexOf(header);
    if (index !== -1) sheet.getRange(row, index + 1).setValue(updates[header]);
  });
  return true;
}

function saveDataUrlFile(folder, dataUrl, name) {
  const match = String(dataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Invalid artifact data URL.');
  const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], name);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getSpeakingArtifactFolder() {
  const folders = DriveApp.getFoldersByName(SPEAKING_ARTIFACT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(SPEAKING_ARTIFACT_FOLDER_NAME);
}

function cleanupOldSpeakingFiles() {
  cleanupOldSpeakingAudioFiles();
  cleanupOldSpeakingArtifacts();
}

function cleanupOldSpeakingArtifacts() {
  try {
    const cutoff = Date.now() - artifactRetentionMs();
    const folders = DriveApp.getFoldersByName(SPEAKING_ARTIFACT_FOLDER_NAME);
    while (folders.hasNext()) {
      const files = folders.next().getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getDateCreated().getTime() < cutoff) file.setTrashed(true);
      }
    }
  } catch (error) {
    console.warn(`Failed to cleanup old speaking artifacts: ${error}`);
  }
}

function installSpeakingCleanupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === 'cleanupOldSpeakingFiles')
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('cleanupOldSpeakingFiles').timeBased().everyDays(1).atHour(3).create();
  return {
    ok: true,
    audioRetentionDays: SPEAKING_AUDIO_RETENTION_DAYS,
    artifactRetentionDays: SPEAKING_ARTIFACT_RETENTION_DAYS
  };
}

function audioRetentionMs() {
  return SPEAKING_AUDIO_RETENTION_DAYS * 24 * 60 * 60 * 1000;
}

function artifactRetentionMs() {
  return SPEAKING_ARTIFACT_RETENTION_DAYS * 24 * 60 * 60 * 1000;
}

function sanitizeFilename(value) {
  return String(value || 'peserta')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || 'peserta';
}

function parseModelJson(text) {
  const raw = String(text || '').trim();
  try {
    return JSON.parse(raw);
  } catch (error) {
    const match = raw.match(/\{[\s\S]*\}/);
    if (match) return JSON.parse(match[0]);
    throw error;
  }
}

function isRetryableKieError(error) {
  const status = Number(error.status || 0);
  const body = String(error.body || error.message || '').toLowerCase();
  return status === 401 ||
    status === 403 ||
    status === 429 ||
    status >= 500 ||
    body.includes('quota') ||
    body.includes('rate limit') ||
    body.includes('resource_exhausted');
}

function buildSpeakingResultFromTasks(taskScores) {
  const totalTasks = Math.max(taskScores.length, 1);
  const maxCriteria = {
    taskResponse: 10,
    fluencyCoherence: 25,
    pronunciation: 25,
    grammar: 20,
    vocabulary: 20
  };
  const criteria = { communicationStrategy: 0 };
  Object.keys(maxCriteria).forEach((key) => {
    const averagePercent = taskScores.reduce((sum, task) => sum + clamp(task.criteriaPercent?.[key], 0, 100), 0) / totalTasks;
    criteria[key] = Math.round((averagePercent / 100) * maxCriteria[key]);
  });
  const overallScore = Math.round(
    criteria.taskResponse +
    criteria.fluencyCoherence +
    criteria.pronunciation +
    criteria.grammar +
    criteria.vocabulary
  );
  const nonZeroTasks = taskScores.filter((task) => !task.isZero && task.score > 0);
  const weakest = getWeakestCriteria(criteria);
  const strongest = getStrongestCriteria(criteria);
  return normalizeSpeakingScore({
    overallScore,
    cefrLevel: getLevelName(overallScore),
    criteria,
    taskScores,
    strengths: nonZeroTasks.length
      ? [`Bagian yang paling kebaca saat ini adalah ${strongest.label}. Ini bisa jadi modal awal buat latihan berikutnya.`]
      : [],
    improvements: [`Next, fokus poles ${weakest.label} dulu biar jawaban kamu makin kebaca dan natural.`],
    teacherNote: nonZeroTasks.length
      ? `Kamu sudah submit ${nonZeroTasks.length} dari ${taskScores.length} audio yang bisa dinilai. Mulai dari feedback per soal, lalu ulangi bagian yang nilainya paling rendah.`
      : 'Tidak ada audio speaking yang bisa dinilai, jadi skor speaking adalah 0.',
    zeroTasks: taskScores.filter((task) => task.isZero || task.score === 0).length
  });
}

function normalizeTaskScore(sourceTask, score) {
  const criteriaPercent = score.criteriaPercent || {};
  const normalized = {
    id: String(score.id || sourceTask.id || ''),
    score: clamp(score.score, 0, 20),
    transcript: String(score.transcript || ''),
    note: String(score.note || ''),
    good: String(score.good || ''),
    correction: String(score.correction || ''),
    grammarFix: String(score.grammarFix || ''),
    pronunciationFix: String(score.pronunciationFix || ''),
    betterAnswer: String(score.betterAnswer || ''),
    scoreReason: String(score.scoreReason || ''),
    isZero: Boolean(score.isZero) || clamp(score.score, 0, 20) === 0,
    criteriaPercent: {
      taskResponse: clamp(criteriaPercent.taskResponse, 0, 100),
      fluencyCoherence: clamp(criteriaPercent.fluencyCoherence, 0, 100),
      pronunciation: clamp(criteriaPercent.pronunciation, 0, 100),
      grammar: clamp(criteriaPercent.grammar, 0, 100),
      vocabulary: clamp(criteriaPercent.vocabulary, 0, 100),
      communicationStrategy: clamp(criteriaPercent.communicationStrategy, 0, 100)
    }
  };
  return applyStrictPlacementCaps(sourceTask, normalized);
}

function applyStrictPlacementCaps(sourceTask, task) {
  if (!isProductiveSpeakingTask(sourceTask.id || task.id) || task.isZero) return task;
  const words = wordCount(task.transcript);
  const criteria = task.criteriaPercent || {};
  const taskResponse = clamp(criteria.taskResponse, 0, 100);
  const grammar = clamp(criteria.grammar, 0, 100);
  const vocabulary = clamp(criteria.vocabulary, 0, 100);
  let cap = 100;
  const reasons = [];

  if (words <= 3) {
    cap = Math.min(cap, 20);
    reasons.push('jawaban hanya berisi beberapa kata');
  } else if (words < 8) {
    cap = Math.min(cap, 35);
    reasons.push('jawaban masih satu kalimat sangat pendek');
  } else if (words < 16) {
    cap = Math.min(cap, 45);
    reasons.push('jawaban masih terlalu singkat dan detailnya belum cukup');
  } else if (words < 25) {
    cap = Math.min(cap, 55);
    reasons.push('jawaban masih pendek untuk menunjukkan kemampuan speaking yang stabil');
  }

  if (taskResponse < 20) {
    cap = Math.min(cap, 25);
    reasons.push('jawaban belum nyambung dengan instruksi utama');
  } else if (taskResponse < 30) {
    cap = Math.min(cap, 35);
    reasons.push('kesesuaian jawaban dengan soal masih lemah');
  } else if (taskResponse < 40) {
    cap = Math.min(cap, 45);
    reasons.push('isi jawaban belum cukup relevan');
  } else if (taskResponse < 60) {
    cap = Math.min(cap, 60);
    reasons.push('jawaban sudah ada, tapi belum cukup lengkap/relevan untuk skor tinggi');
  }

  if (words < 25 && (grammar < 45 || vocabulary < 45)) {
    cap = Math.min(cap, 45);
    reasons.push('jawaban pendek dengan grammar/vocabulary yang masih sangat basic');
  }

  const displayed = clamp(task.score, 0, 20) * 5;
  if (displayed <= cap) return task;

  const ratio = cap / displayed;
  task.score = Math.round(cap / 5);
  Object.keys(task.criteriaPercent).forEach((key) => {
    task.criteriaPercent[key] = Math.round(clamp(task.criteriaPercent[key], 0, 100) * ratio);
  });
  const capNote = `Skor dibatasi di ${cap}/100 karena ${reasons.slice(0, 2).join(' dan ')}.`;
  task.correction = [task.correction, capNote].filter(Boolean).join(' ');
  task.scoreReason = [task.scoreReason, capNote].filter(Boolean).join(' ');
  return task;
}

function isProductiveSpeakingTask(taskId) {
  return !['SPK-001', 'SPK-002'].includes(String(taskId || ''));
}

function wordCount(text) {
  return String(text || '').trim().split(/\s+/).filter((word) => /[A-Za-z]/.test(word)).length;
}

function zeroTaskScore(task, note) {
  return {
    id: task.id,
    score: 0,
    transcript: '',
    note,
    good: '',
    correction: note,
    grammarFix: '',
    pronunciationFix: '',
    betterAnswer: '',
    scoreReason: 'Nilai 0 karena audio tidak tersedia atau belum cukup jelas untuk dinilai.',
    isZero: true,
    criteriaPercent: {
      taskResponse: 0,
      fluencyCoherence: 0,
      pronunciation: 0,
      grammar: 0,
      vocabulary: 0,
      communicationStrategy: 0
    }
  };
}

function getStrongestCriteria(criteria) {
  return criteriaEntry(criteria, (a, b) => b.value - a.value);
}

function getWeakestCriteria(criteria) {
  return criteriaEntry(criteria, (a, b) => a.value - b.value);
}

function criteriaEntry(criteria, sortFn) {
  const labels = {
    taskResponse: 'content',
    fluencyCoherence: 'fluency',
    pronunciation: 'pronunciation',
    grammar: 'grammar',
    vocabulary: 'vocabulary'
  };
  const entries = Object.keys(labels).map((key) => ({ key, label: labels[key], value: Number(criteria[key] || 0) }));
  return entries.sort(sortFn)[0] || { key: 'taskResponse', label: 'content', value: 0 };
}

function normalizeSpeakingScore(score) {
  const criteria = score.criteria || {};
  const normalized = {
    overallScore: clamp(score.overallScore, 0, 100),
    cefrLevel: String(score.cefrLevel || getLevelName(score.overallScore || 0)),
    criteria: {
      taskResponse: clamp(criteria.taskResponse, 0, 10),
      fluencyCoherence: clamp(criteria.fluencyCoherence, 0, 25),
      pronunciation: clamp(criteria.pronunciation, 0, 25),
      grammar: clamp(criteria.grammar, 0, 20),
      vocabulary: clamp(criteria.vocabulary, 0, 20),
      communicationStrategy: 0
    },
    taskScores: Array.isArray(score.taskScores) ? score.taskScores.map((task) => normalizeTaskScore({ id: task.id }, task)) : [],
    strengths: Array.isArray(score.strengths) ? score.strengths.slice(0, 3) : [],
    improvements: Array.isArray(score.improvements) ? score.improvements.slice(0, 3) : [],
    teacherNote: String(score.teacherNote || ''),
    zeroTasks: clamp(score.zeroTasks, 0, 6),
    source: 'gemini'
  };
  normalized.overallScore = Math.round(
    normalized.criteria.taskResponse +
    normalized.criteria.fluencyCoherence +
    normalized.criteria.pronunciation +
    normalized.criteria.grammar +
    normalized.criteria.vocabulary
  );
  normalized.cefrLevel = getLevelName(normalized.overallScore);
  if (normalized.zeroTasks >= 6 || normalized.overallScore === 0) {
    normalized.overallScore = 0;
    normalized.cefrLevel = 'Beginner';
    normalized.teacherNote = normalized.teacherNote || 'Tidak ada audio speaking yang bisa dinilai, jadi skor speaking adalah 0.';
  }
  return normalized;
}

function zeroResult(source) {
  return {
    overallScore: 0,
    cefrLevel: 'Beginner',
    criteria: {
      taskResponse: 0,
      fluencyCoherence: 0,
      pronunciation: 0,
      grammar: 0,
      vocabulary: 0,
      communicationStrategy: 0
    },
    taskScores: [],
    strengths: [],
    improvements: ['Submit clear spoken English audio for each task.'],
    teacherNote: 'Tidak ada audio speaking yang bisa dinilai, jadi skor speaking adalah 0.',
    zeroTasks: 6,
    source
  };
}

function saveSpeakingResult(payload) {
  const sheet = getSheet();
  const user = payload.user || {};
  const result = payload.result || {};
  const criteria = result.criteria || {};
  const artifacts = result.artifacts || {};
  sheet.appendRow([
    new Date(),
    user.name || '',
    user.whatsapp || '',
    user.instagram || '',
    user.instagramFollowers || '',
    user.age || '',
    user.status || '',
    result.overallScore || 0,
    result.cefrLevel || '',
    criteria.taskResponse || 0,
    criteria.fluencyCoherence || 0,
    criteria.pronunciation || 0,
    criteria.grammar || 0,
    criteria.vocabulary || 0,
    criteria.communicationStrategy || 0,
    result.zeroTasks || 0,
    result.teacherNote || '',
    artifacts.certificateUrl || '',
    artifacts.fullReportPdfUrl || '',
    artifacts.expiresAt || ''
  ]);
  return { rowNumber: sheet.getLastRow() };
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  migrateSpeakingSheetHeaders(sheet);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function migrateSpeakingSheetHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(String);
  const emailIndex = headers.indexOf('email');
  const socialIndex = headers.indexOf('social_media_account');
  const followersIndex = headers.indexOf('followers_range');
  const needsFollowersColumn = followersIndex === -1 && (emailIndex !== -1 || socialIndex !== -1);

  if (needsFollowersColumn) {
    const accountColumn = (socialIndex !== -1 ? socialIndex : emailIndex) + 1;
    sheet.insertColumnAfter(accountColumn);
  }
}

function syncSpeakingHeaders() {
  const sheet = getSheet();
  return sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
}

function debugAiProviders() {
  return [{ provider: 'kie.ai gemini-3-5-flash-openai', keys: getKieApiKeys().length }];
}

function getLevelName(score) {
  const value = Number(score || 0);
  if (value <= 25) return 'Beginner';
  if (value <= 45) return 'Elementary';
  if (value <= 60) return 'Pre-Intermediate';
  if (value <= 75) return 'Intermediate';
  if (value <= 88) return 'Upper-Intermediate';
  return 'Advanced';
}

function clamp(value, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return min;
  return Math.max(min, Math.min(max, number));
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
