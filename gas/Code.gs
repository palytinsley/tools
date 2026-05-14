/*
// CLASP SETUP:
// 1. npm install -g @google/clasp
// 2. clasp login
// 3. In repo root: create .clasp.json with scriptId: "1ONMn_IhBqu7ZWu5W1MP-pqrj6ND-q_4jQi2A4qttqhUmSkLl627NAvuo"
// 4. clasp push (from /gas folder)
// 5. In Apps Script editor: Deploy > New deployment > Web app
//    Execute as: Me | Who has access: Anyone
// 6. Copy deployment URL into index.html GAS_URL constant
*/

var SPREADSHEET_ID = '1c_qbf_b3WkAdFbJ0eqeB-qxb70QZPk2m9SHkr_DQrtM';
var STUDENTS_SHEET = 'Students';
var GROUPS_SHEET = 'Groups';
var GRADES_SHEET = 'Grades';

var GRADE_COLUMNS = [
  'Name',
  'Group#',
  'Room',
  'Email',
  'Period',
  'Skip',
  'Winner',
  'Doc Sub /50',
  'Pres Sub /30',
  'Total /80',
  'Pct',
  'Ideas Score',
  'Ideas Comment',
  'Quality Score',
  'Quality Comment',
  'Narration Score',
  'Narration Comment',
  'Video & Images Score',
  'Video & Images Comment',
  'Student Expectations Score',
  'Student Expectations Comment',
  'Day-of Prep Score',
  'Day-of Prep Comment',
  'Preparation Score',
  'Preparation Comment',
  'Content Score',
  'Content Comment',
  'Vocal Skills Score',
  'Vocal Skills Comment',
  'Physical Technique Score',
  'Physical Technique Comment'
];

var GROUP_CRITERIA = [
  { key: 'ideas', scoreHeader: 'Ideas Score', commentHeader: 'Ideas Comment', max: 10, section: 'doc' },
  { key: 'quality', scoreHeader: 'Quality Score', commentHeader: 'Quality Comment', max: 10, section: 'doc' },
  { key: 'narration', scoreHeader: 'Narration Score', commentHeader: 'Narration Comment', max: 10, section: 'doc' },
  { key: 'videoImages', scoreHeader: 'Video & Images Score', commentHeader: 'Video & Images Comment', max: 10, section: 'doc' },
  { key: 'dayOfPrep', scoreHeader: 'Day-of Prep Score', commentHeader: 'Day-of Prep Comment', max: 10, section: 'pres' },
  { key: 'preparation', scoreHeader: 'Preparation Score', commentHeader: 'Preparation Comment', max: 5, section: 'pres' },
  { key: 'content', scoreHeader: 'Content Score', commentHeader: 'Content Comment', max: 5, section: 'pres' }
];

var INDIVIDUAL_CRITERIA = [
  { key: 'studentExpectations', scoreHeader: 'Student Expectations Score', commentHeader: 'Student Expectations Comment', max: 10, section: 'doc' },
  { key: 'vocalSkills', scoreHeader: 'Vocal Skills Score', commentHeader: 'Vocal Skills Comment', max: 5, section: 'pres' },
  { key: 'physicalTechnique', scoreHeader: 'Physical Technique Score', commentHeader: 'Physical Technique Comment', max: 5, section: 'pres' }
];

function doGet() {
  return _jsonResponse({ status: 'ok' });
}

function doPost(e) {
  try {
    var payload = _parseRequestBody(e);
    var action = payload.action;
    var result;

    switch (action) {
      case 'getGroups':
        result = _handleGetGroups();
        break;
      case 'getStudents':
        result = _handleGetStudents(payload);
        break;
      case 'saveGrades':
        result = _handleSaveGrades(payload);
        break;
      case 'getAllGrades':
        result = _handleGetAllGrades();
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }

    return _jsonResponse(result);
  } catch (error) {
    return _jsonResponse({
      success: false,
      error: error.message
    });
  }
}

function _handleGetGroups() {
  var rows = _getDataRows(_getSheet(GROUPS_SHEET));
  var groups = rows.map(function(row) {
    return {
      groupNum: String(row['Group#'] || '').trim(),
      label: String(row['Label'] || '').trim(),
      presLink: String(row['Pres Link'] || '').trim(),
      docLink: String(row['Doc Link'] || '').trim()
    };
  }).filter(function(group) {
    return group.groupNum;
  });

  return { groups: groups };
}

function _handleGetStudents(payload) {
  var groupNum = String(payload.groupNum || '').trim();
  if (!groupNum) {
    throw new Error('groupNum is required for getStudents');
  }

  var studentRows = _getDataRows(_getSheet(STUDENTS_SHEET)).filter(function(row) {
    return String(row['Group#'] || '').trim() === groupNum;
  });
  var gradesMap = _getGradesMapByGroup(groupNum);

  var students = studentRows.map(function(row) {
    var name = String(row['Name'] || '').trim();
    return {
      name: name,
      email: String(row['Email'] || '').trim(),
      room: String(row['Room'] || '').trim(),
      period: String(row['Period'] || '').trim(),
      skip: _normalizeFlag(row['Skip']),
      winner: _normalizeFlag(row['Winner']),
      grades: gradesMap[name] || null
    };
  });

  return { students: students };
}

function _handleSaveGrades(payload) {
  var groupNum = String(payload.groupNum || '').trim();
  var groupScores = payload.groupScores || {};
  var studentScores = payload.studentScores || [];

  if (!groupNum) {
    throw new Error('groupNum is required for saveGrades');
  }
  if (!studentScores.length) {
    throw new Error('studentScores must contain at least one student');
  }

  var studentsByName = {};
  _getDataRows(_getSheet(STUDENTS_SHEET)).forEach(function(row) {
    if (String(row['Group#'] || '').trim() === groupNum) {
      studentsByName[String(row['Name'] || '').trim()] = row;
    }
  });

  var rowsToWrite = [];
  studentScores.forEach(function(studentPayload) {
    var studentName = String(studentPayload.name || '').trim();
    var studentRow = studentsByName[studentName];

    if (!studentRow) {
      throw new Error('Student not found in Students tab for group ' + groupNum + ': ' + studentName);
    }

    var mergedScores = _mergeScores(groupScores, studentPayload);
    _validateScores(mergedScores);
    rowsToWrite.push(_buildGradeRow(studentRow, groupNum, mergedScores));
  });

  var sheet = _getSheet(GRADES_SHEET);
  var existingValues = sheet.getDataRange().getValues();
  if (!existingValues.length) {
    sheet.getRange(1, 1, 1, GRADE_COLUMNS.length).setValues([GRADE_COLUMNS]);
    existingValues = [GRADE_COLUMNS];
  }

  var header = existingValues[0];
  var nameIndex = header.indexOf('Name');
  var groupIndex = header.indexOf('Group#');
  var existingRowMap = {};

  for (var i = 1; i < existingValues.length; i += 1) {
    var existingName = String(existingValues[i][nameIndex] || '').trim();
    var existingGroup = String(existingValues[i][groupIndex] || '').trim();
    if (existingName && existingGroup) {
      existingRowMap[existingName + '||' + existingGroup] = i + 1;
    }
  }

  var appendRows = [];
  rowsToWrite.forEach(function(row) {
    var key = String(row[0]).trim() + '||' + String(row[1]).trim();
    var rowNumber = existingRowMap[key];
    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    } else {
      appendRows.push(row);
    }
  });

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, GRADE_COLUMNS.length).setValues(appendRows);
  }

  return {
    success: true,
    rowsWritten: rowsToWrite.length
  };
}

function _handleGetAllGrades() {
  var rows = _getDataRows(_getSheet(GRADES_SHEET));
  return { grades: rows };
}

function _buildGradeRow(studentRow, groupNum, scores) {
  var docSubtotal = _sumScores(scores, GROUP_CRITERIA.concat([INDIVIDUAL_CRITERIA[0]]), 'doc');
  var presSubtotal = _sumScores(scores, GROUP_CRITERIA.concat(INDIVIDUAL_CRITERIA.slice(1)), 'pres');
  var total = docSubtotal + presSubtotal;
  var pct = (total / 80) * 100;

  var gradeRecord = {
    'Name': String(studentRow['Name'] || '').trim(),
    'Group#': groupNum,
    'Room': String(studentRow['Room'] || '').trim(),
    'Email': String(studentRow['Email'] || '').trim(),
    'Period': String(studentRow['Period'] || '').trim(),
    'Skip': _normalizeFlag(studentRow['Skip']),
    'Winner': _normalizeFlag(studentRow['Winner']),
    'Doc Sub /50': docSubtotal,
    'Pres Sub /30': presSubtotal,
    'Total /80': total,
    'Pct': Math.round(pct * 100) / 100
  };

  GROUP_CRITERIA.forEach(function(criterion) {
    gradeRecord[criterion.scoreHeader] = scores[criterion.key].score;
    gradeRecord[criterion.commentHeader] = scores[criterion.key].comment;
  });

  INDIVIDUAL_CRITERIA.forEach(function(criterion) {
    gradeRecord[criterion.scoreHeader] = scores[criterion.key].score;
    gradeRecord[criterion.commentHeader] = scores[criterion.key].comment;
  });

  return GRADE_COLUMNS.map(function(header) {
    return gradeRecord[header] !== undefined ? gradeRecord[header] : '';
  });
}

function _mergeScores(groupScores, studentPayload) {
  var merged = {};

  GROUP_CRITERIA.forEach(function(criterion) {
    merged[criterion.key] = {
      score: _toNumber(groupScores[criterion.key] && groupScores[criterion.key].score),
      comment: _cleanComment(groupScores[criterion.key] && groupScores[criterion.key].comment)
    };
  });

  INDIVIDUAL_CRITERIA.forEach(function(criterion) {
    merged[criterion.key] = {
      score: _toNumber(studentPayload[criterion.key] && studentPayload[criterion.key].score),
      comment: _cleanComment(studentPayload[criterion.key] && studentPayload[criterion.key].comment)
    };
  });

  return merged;
}

function _validateScores(scores) {
  GROUP_CRITERIA.concat(INDIVIDUAL_CRITERIA).forEach(function(criterion) {
    var entry = scores[criterion.key];
    if (entry === undefined || entry.score === '' || isNaN(entry.score)) {
      throw new Error('Missing score for ' + criterion.key);
    }
    if (entry.score < 0 || entry.score > criterion.max) {
      throw new Error('Score out of range for ' + criterion.key + ' (0-' + criterion.max + ')');
    }
  });
}

function _sumScores(scores, criteria, section) {
  return criteria.reduce(function(total, criterion) {
    if (criterion.section !== section) {
      return total;
    }
    return total + Number(scores[criterion.key].score || 0);
  }, 0);
}

function _getGradesMapByGroup(groupNum) {
  var map = {};
  _getDataRows(_getSheet(GRADES_SHEET)).forEach(function(row) {
    if (String(row['Group#'] || '').trim() === groupNum) {
      var name = String(row['Name'] || '').trim();
      map[name] = {
        docSubtotal: Number(row['Doc Sub /50'] || 0),
        presSubtotal: Number(row['Pres Sub /30'] || 0),
        total: Number(row['Total /80'] || 0),
        pct: Number(row['Pct'] || 0),
        ideas: _gradeEntry(row, 'Ideas Score', 'Ideas Comment'),
        quality: _gradeEntry(row, 'Quality Score', 'Quality Comment'),
        narration: _gradeEntry(row, 'Narration Score', 'Narration Comment'),
        videoImages: _gradeEntry(row, 'Video & Images Score', 'Video & Images Comment'),
        studentExpectations: _gradeEntry(row, 'Student Expectations Score', 'Student Expectations Comment'),
        dayOfPrep: _gradeEntry(row, 'Day-of Prep Score', 'Day-of Prep Comment'),
        preparation: _gradeEntry(row, 'Preparation Score', 'Preparation Comment'),
        content: _gradeEntry(row, 'Content Score', 'Content Comment'),
        vocalSkills: _gradeEntry(row, 'Vocal Skills Score', 'Vocal Skills Comment'),
        physicalTechnique: _gradeEntry(row, 'Physical Technique Score', 'Physical Technique Comment')
      };
    }
  });
  return map;
}

function _gradeEntry(row, scoreHeader, commentHeader) {
  return {
    score: row[scoreHeader] === '' ? '' : Number(row[scoreHeader]),
    comment: String(row[commentHeader] || '')
  };
}

function _getDataRows(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values.length) {
    return [];
  }

  var header = values[0];
  var rows = [];

  for (var i = 1; i < values.length; i += 1) {
    var row = {};
    var hasData = false;
    for (var j = 0; j < header.length; j += 1) {
      row[header[j]] = values[i][j];
      if (values[i][j] !== '' && values[i][j] !== null) {
        hasData = true;
      }
    }
    if (hasData) {
      rows.push(row);
    }
  }

  return rows;
}

function _getSheet(name) {
  var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    throw new Error('Required sheet not found: ' + name);
  }
  return sheet;
}

function _parseRequestBody(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing request body');
  }
  return JSON.parse(e.postData.contents);
}

function _jsonResponse(result) {
  // TODO: Add one-time migration helper in migrate.gs for importing legacy Grades rows from CSPAN_ver_14__2_.xlsx.
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function _toNumber(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  return Number(value);
}

function _cleanComment(value) {
  return String(value || '').trim();
}

function _normalizeFlag(value) {
  return String(value || '').trim().toLowerCase() === 'y' ? 'y' : '';
}
