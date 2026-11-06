// ==============================================================================
// ADMISSION WORKFLOW : GOOGLE APPS SCRIPT + GEMINI API + GMAIL
// Event : GDG DevFest 2026
// Speaker : Cyprien Tankeu (Google Developer Expert Google Workspace)
// ==============================================================================

// 1. CONFIGURATION
const GEMINI_API_KEY = "API Key";
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${GEMINI_API_KEY}`;
const EVENT_NAME = "DevFest 2026";

/**
 * 2. MOTEUR IA : Génère le corps de l'email personnalisé avec Gemini
 */
function generateEmailWithAI(name, status, motivation) {
  const prompt = `
    You are an organizer for the "${EVENT_NAME}".
    Write a personalized confirmation email to "${name}" regarding their admission status: "${status}".
    Their original motivation was: "${motivation}".
    
    Guidelines based on status:
    - "Approved": Warmly welcome them, explain that their motivation stood out, and provide check-in instructions.
    - "Waitlist": Be encouraging, explain that spots are limited, and mention they will be notified if a seat opens.
    - "Rejected": Be respectful and polite, thanking them and encouraging them to join future public webinars.
    
    Output ONLY the email body in clear text without a subject line.
  `;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(GEMINI_URL, options);
    const rawText = response.getContentText();
    const statusCode = response.getResponseCode();

    Logger.log("Code HTTP reçu : " + statusCode);
    Logger.log("Contenu brut reçu : " + rawText);

    // Vérifier si le texte est vide avant d'essayer de le parser
    if (!rawText || rawText.trim() === "") {
      Logger.log("Erreur : La réponse de l'API est totalement vide.");
      return null;
    }

    const data = JSON.parse(rawText);

    if (statusCode !== 200) {
      Logger.log("Erreur API : " + JSON.stringify(data.error || data));
      return null;
    }

    if (data.candidates && data.candidates.length > 0 && data.candidates[0].content) {
      return data.candidates[0].content.parts[0].text;
    } else {
      Logger.log("Réponse sans candidat : " + rawText);
      return null;
    }

  } catch (e) {
    Logger.log("Exception levée : " + e.toString());
    return null;
  }
}


/**
 * 3. PROCESSEUR : Parcourt les lignes du Google Sheet et envoie les emails
 */
function processAdmissionStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  // On commence à la ligne 1 pour sauter la ligne d'en-têtes
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];       // Colonne B : Nom complet
    const email = data[i][2];      // Colonne C : Email
    const motivation = data[i][3]; // Colonne D : Motivation
    const status = data[i][4];     // Colonne E : Statut (Approved / Waitlist / Rejected)
    const emailSent = data[i][5];  // Colonne F : Date d'envoi de l'email

    // Traite la ligne si un statut est renseigné et qu'aucun email n'a encore été envoyé
    if (status && !emailSent) {
      Logger.log(`Génération de l'email pour ${name} (${status})...`);
      const emailBody = generateEmailWithAI(name, status, motivation);

      if (emailBody) {
        const subject = `Update regarding your application: ${EVENT_NAME}`;
        
        // Envoi effectif de l'email via Gmail
        GmailApp.sendEmail(email, subject, emailBody);

        // Horodatage dans la colonne F pour éviter les envois en double
        sheet.getRange(i + 1, 6).setValue(new Date());
        Logger.log(`Email envoyé avec succès à ${email}`);
      }
    }
  }
}

/**
 * 4. MENU PERSONNALISÉ : Ajoute un bouton d'action dans l'interface Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Admissions")
    .addItem("Envoyer les e-mails de statut", "processAdmissionStatuses")
    .addToUi();
}
