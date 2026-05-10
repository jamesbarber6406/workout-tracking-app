function handleEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Submission Form") return;

  if (e.range.getA1Notation() === "F10" && e.value === "TRUE") {
    submitWorkout_(e.source);
  }
}

function submitWorkout_(ss) {
  const form = ss.getSheetByName("Submission Form");
  const log = ss.getSheetByName("Log");

  const timestamp = new Date();
  const submissionId = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 10000);

  const name = cleanInput_(form.getRange("B3").getValue());
  const exerciseName = cleanInput_(form.getRange("B4").getValue());
  const sessionNo = form.getRange("B5").getValue() || 1;
  const comments = form.getRange("B6").getValue();

  const manualCategory = cleanInput_(form.getRange("D3").getValue());
  const mins = Number(form.getRange("D4").getValue()) || 0;
  const secs = Number(form.getRange("D5").getValue()) || 0;

  const manualBodyweightUsed = form.getRange("F3").getValue() === true;
  const manualBodyweightRatio = Number(form.getRange("F4").getValue()) || 0;
  const manualMultiplier = Number(form.getRange("F5").getValue()) || 1;

  const person = getPerson_(ss, name);

  let exercise = null;

  if (exerciseName) {
    exercise = getExercise_(ss, exerciseName);
  }

  const finalCategory = exerciseName ? exercise.category : manualCategory;
  const muscleGroup = exerciseName ? exercise.muscleGroup : "";
  const multiplier = exerciseName ? exercise.multiplier : manualMultiplier;
  const bodyweightUsed = exerciseName ? exercise.bodyweightUsed : manualBodyweightUsed;
  const bodyweightRatio = exerciseName ? exercise.bodyweightRatio : manualBodyweightRatio;

  if (!name) {
    form.getRange("F10").setValue(false);
    SpreadsheetApp.getUi().alert("Please select your name before submitting.");
    return;
  }

  if (!finalCategory) {
    form.getRange("F10").setValue(false);
    SpreadsheetApp.getUi().alert("Please select an exercise or category before submitting.");
    return;
  }

  const rowsToAppend = [];

  if (finalCategory === "Cardio") {
    const totalTime = mins + (secs / 60);
    const cardioScore = totalTime ? totalTime * multiplier : "";

    rowsToAppend.push([
      timestamp,
      name,
      finalCategory,
      exerciseName || "",
      muscleGroup || "",
      "",
      "",
      "",
      "",
      "",
      "",
      totalTime || "",
      cardioScore || "",
      comments || "",
      sessionNo,
      submissionId,
      "",
    ""
    ]);
  } else {
    const setData = form.getRange("B8:D22").getValues();

    setData.forEach(row => {
      const setNo = row[0];
      const weightInput = Number(row[1]) || 0;
      const reps = Number(row[2]) || 0;

      if (!setNo && !weightInput && !reps) return;

      let effectiveWeight = weightInput;

      if (bodyweightUsed && person.weight && bodyweightRatio) {
        effectiveWeight += person.weight * bodyweightRatio;
      }

      const estimated1RM = calculateBrzycki_(weightInput, reps, finalCategory);
      const effectiveEstimated1RM = calculateBrzycki_(effectiveWeight, reps, finalCategory);

      let totalWeight = "";

      if (reps) {
        totalWeight = effectiveWeight * reps;
      } else if (effectiveWeight) {
        totalWeight = effectiveWeight;
      }

      let weightScore = "";

      if (totalWeight && person.weight) {
        weightScore = totalWeight * person.genderMultiplier / person.weight;
      }

      rowsToAppend.push([
        timestamp,
        name,
        finalCategory,
        exerciseName || "",
        muscleGroup || "",
        setNo || "",
        weightInput || "",
        effectiveWeight || "",
        reps || "",
        totalWeight || "",
        weightScore || "",
        "",
        "",
        comments || "",
        sessionNo,
        submissionId,
        estimated1RM || "",
        effectiveEstimated1RM || ""
      ]);
    });
  }

  if (rowsToAppend.length > 0) {
    const timestamps = log.getRange("A:A").getValues();
    let lastDataRow = 1;

    for (let i = timestamps.length - 1; i >= 1; i--) {
      if (timestamps[i][0] !== "") {
        lastDataRow = i + 1;
        break;
      }
    }

    const startRow = lastDataRow + 1;
    log.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  clearSubmissionForm_(form);
}

function cleanInput_(value) {
  if (value === null || value === undefined) return "";
  const text = String(value).trim();

  if (
    text === "" ||
    text === "(blank)" ||
    text === "-- Select --" ||
    text === "-- Clear --" ||
    text === "None"
  ) {
    return "";
  }

  return text;
}

function clearSubmissionForm_(form) {
  form.getRange("B3:B6").clearContent();
  form.getRange("D3:D5").clearContent();
  form.getRange("F3:F5").clearContent();
  form.getRange("B8:D22").clearContent();
  form.getRange("F10").setValue(false);
}

function getPerson_(ss, name) {
  const sheet = ss.getSheetByName("People");
  const data = sheet.getRange("A2:E1000").getValues();

  for (const row of data) {
    if (row[0] === name) {
      return {
        weight: Number(row[1]) || 0,
        genderMultiplier: Number(row[4]) || 1
      };
    }
  }

  return {
    weight: 0,
    genderMultiplier: 1
  };
}

function getExercise_(ss, exerciseName) {
  const sheet = ss.getSheetByName("Exercise List");
  const data = sheet.getRange("A2:H1000").getValues();

  for (const row of data) {
    if (row[0] === exerciseName) {
      return {
        name: row[0],
        muscleGroup: row[1],
        category: row[3],
        multiplier: Number(row[5]) || 1,
        bodyweightUsed: row[6] === true || row[6] === "TRUE" || row[6] === "Yes",
        bodyweightRatio: Number(row[7]) || 0
      };
    }
  }

  return {
    name: exerciseName,
    muscleGroup: "",
    category: "",
    multiplier: 1,
    bodyweightUsed: false,
    bodyweightRatio: 0
  };
}

function calculateBrzycki_(weight, reps, category) {
  weight = Number(weight) || 0;
  reps = Number(reps) || 0;

  if (category === "Cardio") return "";
  if (!weight || !reps || weight <= 0) return "";

  if (reps === 1) return weight;

  return weight / (1.0278 - (0.0278 * reps));
}