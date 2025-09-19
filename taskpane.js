Office.onReady(() => {
  console.log("Practice add-in ready");

  // Attach input listener for live validation
  const input = document.getElementById("linkInput");
  if (input) {
    input.addEventListener("input", validateUrlLive);
  }
});

// Regex for validation
const urlPattern = /^(https?:\/\/)[\w.-]+(\.[a-z]{2,})(:[0-9]{1,5})?(\/.*)?$/i;

function validateUrlLive() {
  let val = document.getElementById("linkInput").value.trim();
  const messageEl = document.getElementById("validationMessage");

  if (!val) {
    messageEl.textContent = "";
    return;
  }

  // Auto-prepend https:// for preview
  if (!/^https?:\/\//i.test(val)) {
    val = "https://" + val;
  }

  if (urlPattern.test(val)) {
    messageEl.textContent = "✔ Valid URL";
    messageEl.className = "valid";
  } else {
    messageEl.textContent = "✖ Invalid URL (must start with http:// or https://)";
    messageEl.className = "invalid";
  }
}

window.insertLogoFromInput = function(event) {
  let hyperlink = document.getElementById("linkInput").value.trim();
  const messageEl = document.getElementById("validationMessage");

  if (!hyperlink) {
    alert("Please enter a URL first.");
    if (event) event.completed && event.completed();
    return;
  }

  // Auto-prepend https:// if missing
  if (!/^https?:\/\//i.test(hyperlink)) {
    hyperlink = "https://" + hyperlink;
  }

  // Validate again before inserting
  if (!urlPattern.test(hyperlink)) {
    alert("Invalid URL. Please correct it before inserting.");
    if (event) event.completed && event.completed();
    return;
  }

  const imageUrl = "https://newchigrboy.github.io/practice-outlook-addin/logo.png";

  Office.context.mailbox.item.body.setSelectedDataAsync(
    `<a href="${hyperlink}" target="_blank"><img src="${imageUrl}" style="max-width:200px;"/></a>`,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to insert logo: " + asyncResult.error.message);
        alert("Something went wrong inserting the logo.");
      } else {
        console.log("Logo inserted successfully: " + hyperlink);
        messageEl.textContent = "✔ Inserted successfully!";
        messageEl.className = "valid";
      }
      if (event) event.completed && event.completed();
    }
  );
};


