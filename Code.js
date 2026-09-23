/**
 * Fonction principale qui sert la page HTML ou gère les actions de validation/rejet par e-mail.
 * @param {Object} e Paramètres de requête de l'URL.
 * @return {HtmlOutput} La page HTML correspondante.
 */
function doGet(e) {
  // Gestion des actions de validation/rejet via les boutons d'e-mail
  if (e && e.parameter && e.parameter.action && e.parameter.id) {
    return handleWebAction(e.parameter.action, e.parameter.id);
  }
  
  // Récupère l'email de l'utilisateur actif avec fallback sécurisé
  var email = getActiveUserEmail();
  
  // Chargement du template Index.html
  var template = HtmlService.createTemplateFromFile('Index');
  template.userEmail = email;
  template.logoDataUri = getHolcimLogoDataUri();
  
  return template.evaluate()
    .setTitle('Safety School - Oggaz Plant')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fonction d'exemple pour inclure des fichiers CSS ou JS externes
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Enregistre une demande de PPI (Permis de Pénétrer Individuel)
 * Appelé depuis le client avec google.script.run.savePPIRequest(data)
 */
function savePPIRequest(data) {
  console.log("Demande PPI reçue : " + JSON.stringify(data));
  // Ici vous pouvez ajouter un code pour insérer dans un Google Sheet :
  var sheet = SpreadsheetApp.openById("1J48XaSE80CshqSdTkcxftMRoGA8Nc6fufRYxPLZbIRI").getSheetByName("PPI");
  sheet.appendRow([new Date(), data.nom, data.email, data.societe, data.zone, data.dateDebut, data.dateFin, data.motif, data.safetyChecked.join(', ')]);
  return { success: true, message: "Demande PPI enregistrée avec succès !" };
}

/**
 * Enregistre un visiteur dans le registre
 * Appelé depuis le client avec google.script.run.saveVisitor(data)
 */
function saveVisitor(data) {
  console.log("Visite enregistrée : " + JSON.stringify(data));
  
  try {
    // 1. Enregistrement de la signature sur Google Drive au format PNG
    var folderId = "1ojE-d5nE_MOdpBgzKVLY4tDYs0GYObxs";
    var folder = DriveApp.getFolderById(folderId);
    
    var base64Data = data.signature.split(',')[1];
    var decoded = Utilities.base64Decode(base64Data);
    var timestamp = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd_HHmmss");
    var fileName = "Signature_" + data.nom.replace(/[^a-zA-Z0-9]/g, "_") + "_" + timestamp + ".png";
    var blob = Utilities.newBlob(decoded, "image/png", fileName);
    
    var file = folder.createFile(blob);
    // Rendre le fichier accessible avec le lien
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    var fileId = file.getId();
    var fileDirectUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
    
    // 2. Enregistrement des informations dans Google Sheets
    var spreadsheetId = "1J48XaSE80CshqSdTkcxftMRoGA8Nc6fufRYxPLZbIRI";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("Registre visiteurs");
    
    if (!sheet) {
      throw new Error("La feuille 'Registre visiteurs' n'existe pas.");
    }
    
    var arriveeFormatted = formatDateTimeString(data.arrivee);
    var departFormatted = formatDateTimeString(data.depart);
    
    sheet.appendRow([
      data.nom,
      data.motif,
      data.adresse,
      data.email,
      arriveeFormatted,
      departFormatted,
      data.motifDetails,
      fileUrl
    ]);
    
    // 3. Récupération des emails des responsables depuis la feuille ADM
    var admSheet = ss.getSheetByName("ADM");
    var emailsToSend = [];
    
    if (admSheet) {
      var values = admSheet.getDataRange().getValues();
      var rolesToNotify = ["Résponsable Ecole sécurité", "Résponsable Safety", "Résponsable Sureté"];
      
      for (var i = 1; i < values.length; i++) { // Skip header row
        var role = values[i][0] ? values[i][0].toString().trim() : "";
        var email = values[i][1] ? values[i][1].toString().trim() : "";
        
        if (role && email) {
          var match = rolesToNotify.some(function(targetRole) {
            return cleanString(role) === cleanString(targetRole);
          });
          if (match) {
            emailsToSend.push(email);
          }
        }
      }
    } else {
      console.warn("La feuille 'ADM' n'a pas été trouvée.");
    }
    
    // 4. Génération du Code QR
    var qrContent = "Nom: " + data.nom + "\nArrivée: " + arriveeFormatted + "\nDépart: " + departFormatted + "\nMotif: " + data.motif;
    var qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + encodeURIComponent(qrContent);
    var qrResponse = UrlFetchApp.fetch(qrUrl);
    var qrBlob = qrResponse.getBlob();
    var qrBase64 = Utilities.base64Encode(qrBlob.getBytes());
    var qrDataUri = "data:image/png;base64," + qrBase64;

    // 5. Génération du Badge PDF
    var pdfBlob = generateVisitorPdfBlob(data, arriveeFormatted, departFormatted, data.signature, qrDataUri);

    // 6. Envoi de l'email de bienvenue au visiteur avec PDF et CC
    var visitorSubject = "Bienvenue à l'usine Oggaz - Votre Pass Visiteur";
    var visitorBody = createVisitorWelcomeEmailHtml(data, arriveeFormatted, departFormatted);
    
    var emailArgs = {
      to: data.email,
      subject: visitorSubject,
      htmlBody: visitorBody,
      attachments: [pdfBlob],
      inlineImages: { holcimLogo: getHolcimLogoBlob() }
    };
    
    if (emailsToSend.length > 0) {
      emailArgs.cc = emailsToSend.join(",");
    }
    
    MailApp.sendEmail(emailArgs);
    
    return { success: true, message: "Enregistrement effectué avec succès !" };
    
  } catch (error) {
    console.error("Erreur saveVisitor: " + error.toString());
    return { success: false, message: "Erreur serveur : " + error.toString() };
  }
}

/**
 * Enregistre une demande d'EPI
 * Appelé depuis le client avec google.script.run.saveEPIRequest(data)
 */
function saveEPIRequest(data) {
  console.log("Demande d'EPI reçue : " + JSON.stringify(data));
  
  try {
    var spreadsheetId = "1J48XaSE80CshqSdTkcxftMRoGA8Nc6fufRYxPLZbIRI";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("DemandesEPI");
    
    // Si la feuille n'existe pas, la créer automatiquement avec les en-têtes
    if (!sheet) {
      sheet = ss.insertSheet("DemandesEPI");
      sheet.appendRow([
        "ID Demande",
        "Date Demande",
        "Nom & Prénom",
        "Société",
        "Email",
        "Équipements demandés",
        "Date de retrait",
        "Statut",
        "Notification Envoyée"
      ]);
      sheet.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f1f5f9");
    }
    
    var timestamp = new Date();
    var idDemande = "EPI_" + timestamp.getTime();
    var dateDemandeFormatted = Utilities.formatDate(timestamp, "GMT+1", "dd/MM/yyyy HH:mm");
    var dateRetraitFormatted = formatDateTimeString(data.dateRetrait);
    var equipmentsText = formatEquipmentsText(data.equipments);
    
    sheet.appendRow([
      idDemande,
      dateDemandeFormatted,
      data.nom,
      data.societe,
      data.email,
      equipmentsText,
      dateRetraitFormatted,
      "En attente",
      "Non"
    ]);
    
    // Récupération de l'URL de l'application Web pour les boutons de validation
    var webAppUrl = "";
    try {
      webAppUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      webAppUrl = "https://script.google.com/macros/s/AKfycbx8o6mhcq5vZMArk-xSKGjdLJvsSw2eaClYYLY0E5XyNNmkSe7rV2VYWI15cS8Cbzk1/exec";
    }
    
    if (!webAppUrl || webAppUrl === "") {
      // Fallback au cas où l'URL n'est pas récupérable directement
      webAppUrl = "https://script.google.com/macros/s/AKfycbx8o6mhcq5vZMArk-xSKGjdLJvsSw2eaClYYLY0E5XyNNmkSe7rV2VYWI15cS8Cbzk1/exec";
    }
    
    var approveUrl = webAppUrl + "?action=approve&id=" + idDemande;
    var rejectUrl = webAppUrl + "?action=reject&id=" + idDemande;
    
    // Envoi de l'email de notification de demande d'EPI avec boutons
    var toEmail = "noureddine.bouchentouf@holcim.com";
    var ccEmail = "toufik.rabahi@holcim.com";
    var subject = "Nouvelle demande d'EPI à valider - " + data.nom;
    var htmlBody = createEPINotificationEmailHtml(data, equipmentsText, dateRetraitFormatted, approveUrl, rejectUrl);
    
    MailApp.sendEmail({
      to: toEmail,
      cc: ccEmail,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: { holcimLogo: getHolcimLogoBlob() }
    });
    
    return { success: true, message: "Demande d'EPI enregistrée avec succès !" };
    
  } catch (error) {
    console.error("Erreur saveEPIRequest: " + error.toString());
    return { success: false, message: "Erreur serveur : " + error.toString() };
  }
}

/**
 * Gère l'action de validation ou rejet d'une demande d'EPI via l'application Web
 */
function handleWebAction(action, id) {
  try {
    var spreadsheetId = "1J48XaSE80CshqSdTkcxftMRoGA8Nc6fufRYxPLZbIRI";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName("DemandesEPI");
    
    if (!sheet) {
      return HtmlService.createHtmlOutput("<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #ef4444; border-radius: 8px;'>" +
                                           "<h3 style='color: #ef4444;'>Erreur</h3>" +
                                           "<p>La feuille 'DemandesEPI' est introuvable.</p>" +
                                           "</div>")
        .setTitle("Safety School - Erreur Action")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    
    // Localisation dynamique des colonnes par rapport aux en-têtes
    var idColIdx = headers.indexOf("ID Demande");
    var nomColIdx = headers.indexOf("Nom & Prénom");
    var emailColIdx = headers.indexOf("Email");
    var equipementsColIdx = headers.indexOf("Équipements demandés");
    var dateRetraitColIdx = headers.indexOf("Date de retrait");
    var statutColIdx = headers.indexOf("Statut");
    var notifieColIdx = headers.indexOf("Notification Envoyée");
    
    // Fallbacks par défaut si les en-têtes sont absents
    if (idColIdx === -1) idColIdx = 0;
    if (nomColIdx === -1) nomColIdx = 2;
    if (emailColIdx === -1) emailColIdx = 4;
    if (equipementsColIdx === -1) equipementsColIdx = 5;
    if (dateRetraitColIdx === -1) dateRetraitColIdx = 6;
    if (statutColIdx === -1) statutColIdx = 7;
    if (notifieColIdx === -1) notifieColIdx = 8;
    
    var foundRow = -1;
    var rowData = null;
    
    for (var i = 1; i < values.length; i++) {
      if (values[i][idColIdx] === id) {
        foundRow = i + 1; // 1-based row index
        rowData = values[i];
        break;
      }
    }
    
    if (foundRow === -1) {
      return HtmlService.createHtmlOutput("<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #ef4444; border-radius: 8px;'>" +
                                           "<h3 style='color: #ef4444;'>Erreur</h3>" +
                                           "<p>La demande d'EPI avec l'identifiant " + id + " est introuvable.</p>" +
                                           "</div>")
        .setTitle("Safety School - Erreur Action")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
    
    var nom = rowData[nomColIdx];
    var email = rowData[emailColIdx] ? rowData[emailColIdx].toString().trim() : "";
    var equipements = rowData[equipementsColIdx];
    var dateRetrait = rowData[dateRetraitColIdx];
    var currentStatut = rowData[statutColIdx];
    var notifie = rowData[notifieColIdx];
    
    var htmlResponse = "";
    
    if (action === "approve") {
      if (currentStatut === "Validé") {
        htmlResponse = "<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>" +
                       "<div style='color: #10b981; font-size: 48px; margin-bottom: 15px;'>✓</div>" +
                       "<h2>Demande déjà validée</h2>" +
                       "<p>La demande d'EPI pour <strong>" + escapeHtml(nom) + "</strong> a déjà été validée précédemment.</p>" +
                       "</div>";
      } else {
        // Vérification de la présence de l'e-mail du demandeur
        if (!email || email === "") {
          throw new Error("L'adresse e-mail du demandeur est vide dans la feuille de calcul (colonne '" + headers[emailColIdx] + "' à la ligne " + foundRow + "). Impossible d'envoyer l'e-mail de validation.");
        }
        
        // Mettre à jour le statut dans le Google Sheet
        sheet.getRange(foundRow, statutColIdx + 1).setValue("Validé");
        
        // Envoi de l'email de validation au demandeur
        sendEPIValidationEmail(email, nom, equipements, dateRetrait);
        sheet.getRange(foundRow, notifieColIdx + 1).setValue("Oui");
        
        htmlResponse = "<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>" +
                       "<div style='color: #10b981; font-size: 48px; margin-bottom: 15px;'>✓</div>" +
                       "<h2>Demande Validée avec Succès</h2>" +
                       "<p>La demande d'EPI pour <strong>" + escapeHtml(nom) + "</strong> a été validée. Un e-mail de confirmation a été envoyé à <strong>" + escapeHtml(email) + "</strong>.</p>" +
                       "</div>";
      }
    } else if (action === "reject") {
      if (currentStatut === "Rejeté") {
        htmlResponse = "<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>" +
                       "<div style='color: #ef4444; font-size: 48px; margin-bottom: 15px;'>✗</div>" +
                       "<h2>Demande déjà rejetée</h2>" +
                       "<p>La demande d'EPI pour <strong>" + escapeHtml(nom) + "</strong> a déjà été rejetée précédemment.</p>" +
                       "</div>";
      } else {
        // Vérification de la présence de l'e-mail du demandeur
        if (!email || email === "") {
          throw new Error("L'adresse e-mail du demandeur est vide dans la feuille de calcul (colonne '" + headers[emailColIdx] + "' à la ligne " + foundRow + "). Impossible d'envoyer l'e-mail de rejet.");
        }
        
        // Mettre à jour le statut dans le Google Sheet
        sheet.getRange(foundRow, statutColIdx + 1).setValue("Rejeté");
        
        // Envoi de l'email de rejet au demandeur
        sendEPIRejectionEmail(email, nom, equipements, dateRetrait);
        sheet.getRange(foundRow, notifieColIdx + 1).setValue("Rejeté");
        
        htmlResponse = "<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);'>" +
                       "<div style='color: #ef4444; font-size: 48px; margin-bottom: 15px;'>✗</div>" +
                       "<h2>Demande Rejetée</h2>" +
                       "<p>La demande d'EPI pour <strong>" + escapeHtml(nom) + "</strong> a été rejetée. Un e-mail de notification a été envoyé au demandeur.</p>" +
                       "</div>";
      }
    }
    
    return HtmlService.createHtmlOutput(htmlResponse)
      .setTitle("Safety School - Action EPI")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch (error) {
    return HtmlService.createHtmlOutput("<div style='font-family: Arial, sans-serif; max-width: 500px; margin: 50px auto; text-align: center; padding: 30px; border: 1px solid #ef4444; border-radius: 8px;'>" +
                                         "<h3 style='color: #ef4444;'>Erreur lors de la validation</h3>" +
                                         "<p>" + escapeHtml(error.toString()) + "</p>" +
                                         "</div>")
      .setTitle("Safety School - Erreur Action")
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// ------------------ FONCTIONS UTILITAIRES / UTILS ------------------

function formatDateTimeString(dateTimeStr) {
  if (!dateTimeStr) return "";
  try {
    var date = new Date(dateTimeStr.replace(' ', 'T'));
    if (isNaN(date.getTime())) return dateTimeStr;
    return Utilities.formatDate(date, "GMT+1", "dd/MM/yyyy HH:mm");
  } catch (e) {
    return dateTimeStr;
  }
}

function cleanString(str) {
  if (!str) return "";
  return str.toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
            .replace(/\s+/g, " ")
            .trim();
}

function formatEquipmentsText(equipments) {
  if (!equipments || !Array.isArray(equipments)) return "";
  return equipments.map(function(eq) {
    if (eq.size) {
      return eq.item + " (Taille: " + eq.size + ")";
    }
    return eq.item;
  }).join(", ");
}

function createEmailHtml(data, fileUrl, fileDirectUrl) {
  var arriveeFormatted = formatDateTimeString(data.arrivee);
  var departFormatted = formatDateTimeString(data.depart);
  
  return '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">' +
         '<div style="background-color: #003366; color: white; padding: 15px; border-radius: 6px 6px 0 0; text-align: center; margin-bottom: 20px;">' +
         '<h2 style="margin: 0; font-size: 20px;">Safety School - Usine Oggaz</h2>' +
         '<p style="margin: 5px 0 0 0; font-size: 14px;">Notification de nouveau visiteur</p>' +
         '</div>' +
         '<p>Bonjour,</p>' +
         '<p>Un nouveau visiteur a complété sa formation de sécurité et s\'est enregistré dans le registre des visiteurs :</p>' +
         '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold; width: 40%;">Nom & Prénom :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.nom) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Motif de la visite :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.motif) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Adresse :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.adresse) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Email :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.email) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Date & Heure d\'arrivée :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + arriveeFormatted + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Date & Heure de départ prévues :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + departFormatted + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Précisions sur la visite :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.motifDetails || "Aucune précision") + '</td></tr>' +
         '</table>' +
         '<div style="margin-top: 25px; border-top: 1px solid #e2e8f0; padding-top: 15px;">' +
         '<p style="font-weight: bold; margin-bottom: 10px;">Signature Digitale :</p>' +
         '<a href="' + fileUrl + '" target="_blank" style="display: inline-block; background-color: #e30613; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 14px; margin-bottom: 15px;">Voir l\'image de la signature</a>' +
         '<br><img src="' + fileDirectUrl + '" alt="Signature" style="max-width: 250px; border: 1px solid #cbd5e1; border-radius: 6px; padding: 5px; max-height: 120px;" />' +
         '</div>' +
         '<hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />' +
         '<p style="font-size: 11px; color: #718096; text-align: center;">Ceci est un message automatique envoyé par le système Safety School de l\'usine Holcim Oggaz.</p>' +
         '</div>';
}

function createEPINotificationEmailHtml(data, equipmentsText, dateRetraitFormatted, approveUrl, rejectUrl) {
  return '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">' +
         '<div style="background-color: #003366; color: white; padding: 15px; border-radius: 6px 6px 0 0; text-align: center; margin-bottom: 20px;">' +
         '<img src="cid:holcimLogo" style="max-width: 180px; max-height: 55px; object-fit: contain; background: white; padding: 5px; border-radius: 4px;" alt="Logo Holcim" />' +
         '<h2 style="margin: 0; font-size: 20px;">Safety School - Usine Oggaz</h2>' +
         '<p style="margin: 5px 0 0 0; font-size: 14px;">Nouvelle demande d\'EPI à valider</p>' +
         '</div>' +
         '<p>Bonjour,</p>' +
         '<p>Une nouvelle demande d\'équipements de protection individuelle (EPI) a été enregistrée :</p>' +
         '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold; width: 40%;">Demandeur :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.nom) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Société :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.societe) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Email :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + escapeHtml(data.email) + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Date de retrait prévue :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + dateRetraitFormatted + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Équipements demandés :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7; color: #e30613; font-weight: bold;">' + escapeHtml(equipmentsText) + '</td></tr>' +
         '</table>' +
         '<p style="margin-top: 25px; font-weight: bold; text-align: center;">Veuillez cliquer sur un bouton ci-dessous pour traiter cette demande :</p>' +
         '<div style="text-align: center; margin: 20px 0; font-size: 14px;">' +
         '<a href="' + approveUrl + '" style="display: inline-block; background-color: #10b981; color: white; padding: 12px 25px; text-decoration: none; border-radius: 6px; font-weight: bold; margin: 0 10px;">✓ Validé</a>' +
         '<a href="' + rejectUrl + '" style="display: inline-block; background-color: #ef4444; color: white; padding: 12px 25px; text-decoration: none; border-radius: 6px; font-weight: bold; margin: 0 10px;">✗ Rejeté</a>' +
         '</div>' +
         '<hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />' +
         '<p style="font-size: 11px; color: #718096; text-align: center;">Ceci est un message automatique envoyé par le système Safety School de l\'usine Holcim Oggaz.</p>' +
         '</div>';
}

function sendEPIValidationEmail(email, nom, equipements, dateRetrait) {
  var subject = "Votre demande d'EPI a été validée - Safety School Oggaz";
  var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">' +
         '<div style="background-color: #10b981; color: white; padding: 15px; border-radius: 6px 6px 0 0; text-align: center; margin-bottom: 20px;">' +
      '<img src="cid:holcimLogo" style="max-width: 180px; max-height: 55px; object-fit: contain; background: white; padding: 5px; border-radius: 4px;" alt="Logo Holcim" />' +
         '<h2 style="margin: 0; font-size: 20px;">Safety School - Usine Oggaz</h2>' +
         '<p style="margin: 5px 0 0 0; font-size: 14px;">Confirmation de validation d\'EPI</p>' +
         '</div>' +
         '<p>Bonjour ' + escapeHtml(nom) + ',</p>' +
         '<p>Nous avons le plaisir de vous informer que votre demande d\'équipements de protection individuelle (EPI) a été <strong>validée</strong> par les responsables HSE :</p>' +
         '<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold; width: 40%;">Date de retrait prévue :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7;">' + dateRetrait + '</td></tr>' +
         '<tr><td style="padding: 8px; border-bottom: 1px solid #edf2f7; font-weight: bold;">Équipements validés :</td><td style="padding: 8px; border-bottom: 1px solid #edf2f7; color: #047857; font-weight: bold;">' + escapeHtml(equipements) + '</td></tr>' +
         '</table>' +
         '<p>Vous pouvez récupérer ces équipements au bureau de sécurité dès votre arrivée sur site le jour de votre visite.</p>' +
         '<p>Nous vous rappelons que le port des EPI de base est obligatoire dans toutes les zones rouges de l\'usine.</p>' +
         '<p>Bonne visite et travaillez en toute sécurité !</p>' +
         '<hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />' +
         '<p style="font-size: 11px; color: #718096; text-align: center;">Ceci est un message automatique envoyé par le système Safety School de l\'usine Holcim Oggaz.</p>' +
         '</div>';
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: { holcimLogo: getHolcimLogoBlob() }
  });
}

function sendEPIRejectionEmail(email, nom, equipements, dateRetrait) {
  var subject = "Votre demande d'EPI a été rejetée - Safety School Oggaz";
  var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">' +
         '<div style="background-color: #ef4444; color: white; padding: 15px; border-radius: 6px 6px 0 0; text-align: center; margin-bottom: 20px;">' +
      '<img src="cid:holcimLogo" style="max-width: 180px; max-height: 55px; object-fit: contain; background: white; padding: 5px; border-radius: 4px;" alt="Logo Holcim" />' +
         '<h2 style="margin: 0; font-size: 20px;">Safety School - Usine Oggaz</h2>' +
         '<p style="margin: 5px 0 0 0; font-size: 14px;">Notification de rejet de demande d\'EPI</p>' +
         '</div>' +
         '<p>Bonjour ' + escapeHtml(nom) + ',</p>' +
         '<p>Nous vous informons que votre demande d\'équipements de protection individuelle (EPI) pour le ' + dateRetrait + ' a été <strong>rejetée</strong> par l\'équipe de sécurité.</p>' +
         '<p><strong>Détails des équipements :</strong> ' + escapeHtml(equipements) + '</p>' +
         '<p>Si vous avez des questions ou souhaitez soumettre une nouvelle demande, veuillez contacter l\'équipe de sécurité du site.</p>' +
         '<hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 20px 0;" />' +
         '<p style="font-size: 11px; color: #718096; text-align: center;">Ceci est un message automatique envoyé par le système Safety School de l\'usine Holcim Oggaz.</p>' +
         '</div>';
  
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    inlineImages: { holcimLogo: getHolcimLogoBlob() }
  });
}

/**
 * Génère le blob PDF du badge visiteur à partir d'un template HTML.
 */
function generateVisitorPdfBlob(data, arriveeFormatted, departFormatted, signatureDataUri, qrDataUri) {
  var logoBlob = getHolcimLogoBlob();
  var logoDataUri = "data:" + logoBlob.getContentType() + ";base64," + Utilities.base64Encode(logoBlob.getBytes());

  var html = '<div style="font-family: Arial, sans-serif; width: 100%; max-width: 800px; margin: 0 auto; border: 2px solid #003366; padding: 20px;">' +
             '<div style="text-align: center; border-bottom: 2px solid #e30613; padding-bottom: 15px; margin-bottom: 20px;">' +
             '<img src="' + logoDataUri + '" style="max-width: 220px; max-height: 70px; object-fit: contain; margin-bottom: 10px;" alt="Logo Holcim" />' +
             '<h1 style="color: #003366; margin: 0;">PASS VISITEUR - USINE OGGAZ</h1>' +
             '<h3 style="color: #666; margin: 5px 0 0 0;">Système Safety School</h3>' +
             '</div>' +
             '<table style="width: 100%; border-collapse: collapse;">' +
             '<tr>' +
             '<td style="width: 70%; padding-right: 20px; vertical-align: top;">' +
             '<h2 style="color: #003366; border-bottom: 1px solid #ccc; padding-bottom: 5px;">Informations du Visiteur</h2>' +
             '<p><strong>Nom & Prénom :</strong> ' + escapeHtml(data.nom) + '</p>' +
             '<p><strong>Motif de la visite :</strong> ' + escapeHtml(data.motif) + '</p>' +
             '<p><strong>Adresse :</strong> ' + escapeHtml(data.adresse) + '</p>' +
             '<p><strong>Date & Heure d\'arrivée :</strong> ' + arriveeFormatted + '</p>' +
             '<p><strong>Date & Heure de départ prévues :</strong> ' + departFormatted + '</p>' +
             '<p><strong>Précisions sur la visite :</strong> ' + escapeHtml(data.motifDetails || 'N/A') + '</p>' +
             '</td>' +
             '<td style="width: 30%; text-align: center; vertical-align: top; border-left: 1px solid #ccc; padding-left: 20px;">' +
             '<h3 style="color: #003366; margin-top: 0;">Code QR d\'Accès</h3>' +
             '<img src="' + qrDataUri + '" style="width: 150px; height: 150px; border: 1px solid #ccc; padding: 5px;" alt="QR Code" />' +
             '</td>' +
             '</tr>' +
             '</table>' +
             '<div style="margin-top: 30px; border-top: 1px solid #ccc; padding-top: 15px;">' +
             '<h3 style="color: #003366; margin-top: 0;">Signature du Visiteur</h3>' +
             '<img src="' + signatureDataUri + '" style="max-height: 80px; max-width: 300px; border: 1px dashed #ccc; padding: 5px;" alt="Signature" />' +
             '<p style="font-size: 10px; color: #999;">Document généré automatiquement.</p>' +
             '</div>' +
             '</div>';
             
  var blob = Utilities.newBlob(html, MimeType.HTML, "badge_visiteur.html");
  var pdfBlob = blob.getAs(MimeType.PDF);
  pdfBlob.setName("Badge_Visiteur_" + data.nom.replace(/[^a-zA-Z0-9]/g, "_") + ".pdf");
  
  return pdfBlob;
}

/**
 * Génère le contenu HTML de l'email de bienvenue pour le visiteur.
 */
function createVisitorWelcomeEmailHtml(data, arriveeFormatted, departFormatted) {
  return '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">' +
         '<div style="background-color: #003366; color: white; padding: 15px; border-radius: 6px 6px 0 0; text-align: center; margin-bottom: 20px;">' +
         '<img src="cid:holcimLogo" style="max-width: 180px; max-height: 55px; object-fit: contain; background: white; padding: 5px; border-radius: 4px;" alt="Logo Holcim" />' +
         '<h2 style="margin: 0; font-size: 20px;">Safety School - Usine Oggaz</h2>' +
         '<p style="margin: 5px 0 0 0; font-size: 14px;">Bienvenue sur notre site</p>' +
         '</div>' +
         '<p>Bonjour <strong>' + escapeHtml(data.nom) + '</strong>,</p>' +
         '<p>Nous avons le plaisir de vous confirmer votre enregistrement dans le registre de sécurité de l\'<strong>usine Holcim Oggaz</strong>.</p>' +
         '<p>Vous trouverez ci-joint votre <strong>Pass Visiteur PDF</strong> contenant un code QR et vos informations d\'accès. Veuillez présenter ce pass à l\'accueil ou aux agents de sécurité lors de votre arrivée sur le site.</p>' +
         '<table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #f8fafc; border-radius: 6px; border: 1px solid #e2e8f0;">' +
         '<tr><td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; width: 40%; color: #003366;">Nom & Prénom :</td><td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">' + escapeHtml(data.nom) + '</td></tr>' +
         '<tr><td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #003366;">Motif de la visite :</td><td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">' + escapeHtml(data.motif) + '</td></tr>' +
         '<tr><td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #003366;">Adresse :</td><td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">' + escapeHtml(data.adresse) + '</td></tr>' +
         '<tr><td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #003366;">Date & Heure d\'arrivée :</td><td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">' + arriveeFormatted + '</td></tr>' +
         '<tr><td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-weight: bold; color: #003366;">Date & Heure de départ :</td><td style="padding: 10px; border-bottom: 1px solid #e2e8f0;">' + departFormatted + '</td></tr>' +
         '</table>' +
         '<p>Le port des équipements de protection individuelle (EPI) de base est strictement obligatoire sur le site.</p>' +
         '<p>Nous vous souhaitons une excellente visite. Travaillez en toute sécurité !</p>' +
         '<p style="font-size: 11px; color: #718096; text-align: center;">Ceci est un message automatique envoyé par le système Safety School de l\'usine Holcim Oggaz.</p>' +
         '</div>';
}

/**
 * Récupère dynamiquement l'e-mail de l'utilisateur actif pour la traçabilité.
 * Utilisé par le frontend si l'injection de template échoue.
 */
function getActiveUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function getCurrentUserInfo() {
  var email = getActiveUserEmail();

  var webAppUrl = "";
  try {
    webAppUrl = ScriptApp.getService().getUrl();
  } catch (e) {
    webAppUrl = "";
  }

  return {
    email: email || "",
    webAppUrl: webAppUrl
  };
}

function getHolcimLogoBlob() {
  var logoFileId = "13i-h47b3q-l9ZQGjr5Y0VaB7xBnqRyPG";
  return DriveApp.getFileById(logoFileId).getBlob().setName("Holcim-Logo");
}

function getHolcimLogoDataUri() {
  var logoBlob = getHolcimLogoBlob();
  return "data:" + logoBlob.getContentType() + ";base64," + Utilities.base64Encode(logoBlob.getBytes());
}

function escapeHtml(text) {
  if (text === undefined || text === null) return "";
  return text.toString()
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
}