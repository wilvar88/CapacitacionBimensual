/**
 * Código para Google Apps Script (Código.gs)
 * Este script lee las hojas del Google Sheet, calcula los KPIs y devuelve un JSON
 * para que tu Dashboard lo consuma. Además, incluye análisis con Gemini AI.
 */

function doGet(e) {
  // Manejo de CORS preflight request (opcional dependiendo de cómo se consulte)
  return handleResponse();
}

function handleResponse() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  var allCoursesData = {};
  var allAreasData = {};
  var allFeedbacks = [];
  var totalAttendees = 0;
  var globalFinished = 0;
  var globalApproved = 0;
  var globalStarted = 0;
  var globalRatingsCount = 0;
  var globalRatingsSum = 0;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) continue;
    
    var headers = data[0];
    var colIndex = {};
    for (var h = 0; h < headers.length; h++) {
       if(headers[h]) colIndex[headers[h].toString().trim()] = h;
    }
    
    if (colIndex['Curso'] === undefined) continue;

    var participantColName = Object.keys(colIndex).find(function(k) { return k.match(/nombre|colaborador|participante|alumno|empleado/i); });
    var areaColName = Object.keys(colIndex).find(function(k) { return k.toLowerCase() === 'área' || k.toLowerCase() === 'area'; });

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var cursoName = row[colIndex['Curso']];
      if (!cursoName) continue;
      
      var areaName = areaColName ? row[colIndex[areaColName]].toString().trim() : sheet.getName();
      if (!areaName) areaName = "Sin Área";

      if (!allAreasData[areaName]) {
        allAreasData[areaName] = {
          name: areaName,
          total_enrolled: 0,
          not_started: 0,
          approved: 0,
          finished: 0,
          participants: []
        };
      }

      var estadoColaborador = colIndex['Estado'] !== undefined ? row[colIndex['Estado']] : "Activo";
      var estadoAprobacion = colIndex['Estado de Aprobación'] !== undefined ? row[colIndex['Estado de Aprobación']].toString().trim() : "";
      var participantName = participantColName ? row[colIndex[participantColName]] : ("Participante " + r);
      
      if (!allCoursesData[cursoName]) {
         allCoursesData[cursoName] = {
            name: cursoName, category: areaName, total_enrolled: 0, not_started: 0, approved: 0, finished: 0, ratings_count: 0, ratings_sum: 0, activos: 0, inactivos: 0
         };
      }
      
      var c = allCoursesData[cursoName];
      var a = allAreasData[areaName];
      
      c.total_enrolled++;
      a.total_enrolled++;
      totalAttendees++;

      if (estadoColaborador === "Activo") { c.activos++; } else { c.inactivos++; }

      var lowerEstado = estadoAprobacion.toLowerCase();

      if (lowerEstado.indexOf("inscrit") !== -1 || lowerEstado === "") {
        c.not_started++;
        a.not_started++;
      } else {
        if (lowerEstado.indexOf("aprobado") !== -1) {
          c.approved++; a.approved++;
          c.finished++; a.finished++;
        } else if (lowerEstado.indexOf("reproba") !== -1 || lowerEstado.indexOf("finaliz") !== -1 || lowerEstado.indexOf("reprobado") !== -1) {
          c.finished++; a.finished++;
        }
      }

      a.participants.push({
         name: participantName,
         course: cursoName,
         status: estadoAprobacion
      });

      var valoracion = colIndex['Valoración'] !== undefined ? row[colIndex['Valoración']] : "";
      if (valoracion !== "" && valoracion != null && !isNaN(valoracion)) {
         c.ratings_count++;
         c.ratings_sum += parseFloat(valoracion);
         globalRatingsCount++;
         globalRatingsSum += parseFloat(valoracion);
         
         var comentario = colIndex['Comentario'] !== undefined ? row[colIndex['Comentario']] : "";
         if (comentario && comentario.toString().length > 3) {
            allFeedbacks.push("Área " + areaName + " | Curso: " + cursoName + " | Valoración: " + valoracion + " | Comentario: " + comentario);
         }
      }
    }
  }

  var coursesOutput = [];
  for (var key in allCoursesData) {
     var c = allCoursesData[key];
     var started = c.total_enrolled - c.not_started;
     c.participation = c.total_enrolled > 0 ? (started / c.total_enrolled) * 100 : 0;
     c.approval = c.total_enrolled > 0 ? (c.approved / c.total_enrolled) * 100 : 0;
     c.feedback_participation = c.total_enrolled > 0 ? (c.ratings_count / c.total_enrolled) * 100 : 0;
     c.average_rating = c.ratings_count > 0 ? (c.ratings_sum / c.ratings_count) : 0;
     c.enrolled = c.total_enrolled; 
     
     globalStarted += started;
     globalApproved += c.approved;
     globalFinished += c.finished;
     coursesOutput.push(c);
  }

  var areasOutput = [];
  for (var key in allAreasData) {
     var a = allAreasData[key];
     var aStarted = a.total_enrolled - a.not_started;
     a.participation = a.total_enrolled > 0 ? (aStarted / a.total_enrolled) * 100 : 0;
     a.approval = a.total_enrolled > 0 ? (a.approved / a.total_enrolled) * 100 : 0;
     a.enrolled = a.total_enrolled;
     areasOutput.push(a);
  }

  var globalAvgRating = globalRatingsCount > 0 ? (globalRatingsSum / globalRatingsCount) : 0;
  var globalRatingParticipation = totalAttendees > 0 ? (globalRatingsCount / totalAttendees) * 100 : 0;

  var lastUpdatedStr = "";
  try {
    var file = DriveApp.getFileById(spreadsheet.getId());
    lastUpdatedStr = Utilities.formatDate(file.getLastUpdated(), "GMT-5", "dd/MM/yyyy HH:mm");
  } catch(e) {
    lastUpdatedStr = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm") + " (Local)";
  }

  // Leer Análisis de Valoración desde la Hoja creada por el usuario
  var geminiAnalysis = { positive: "El administrador no ha registrado el análisis aún.", improvement: "El administrador no ha registrado el análisis aún." };
  var insightsSheet = spreadsheet.getSheetByName("Analisis de Valoración");
  
  if (insightsSheet) {
    var positiveText = insightsSheet.getRange("A2").getValue();
    var improvementText = insightsSheet.getRange("B2").getValue();
    
    if (positiveText || improvementText) {
      geminiAnalysis = {
         positive: positiveText || "No hay comentarios positivos registrados.",
         improvement: improvementText || "No hay oportunidades de mejora registradas."
      };
    }
  }

  var response = {
    kpis: {
      total_courses: coursesOutput.length,
      total_areas: areasOutput.length,
      total_attendees: totalAttendees,
      global_started: globalStarted,
      global_approved: globalApproved,
      participation_rate: totalAttendees > 0 ? (globalStarted / totalAttendees) * 100 : 0,
      approval_rate: totalAttendees > 0 ? (globalApproved / totalAttendees) * 100 : 0,
      average_rating: globalAvgRating,
      ratings_count: globalRatingsCount,
      rating_participation: globalRatingParticipation,
      last_updated: lastUpdatedStr,
      download_url: "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/export?format=xlsx&gid=" + (spreadsheet.getSheets().length > 0 ? spreadsheet.getSheets()[0].getSheetId() : 0)
    },
    courses: coursesOutput,
    areas: areasOutput,
    ai_insights: geminiAnalysis
  };
  
  var jsonParams = JSON.stringify(response);
  return ContentService.createTextOutput(jsonParams).setMimeType(ContentService.MimeType.JSON);
}


