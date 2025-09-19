Office.onReady(() => {
  console.log("Practice add-in ready");
});

window.insertLogoFromInput = function(event) {
  let hyperlink = document.getElementById("linkInput").value.trim();

  // Auto-prepend https:// if missing
  if (hyperlink && !/^https?:\/\//i.test(hyperlink)) {
    hyperlink = "https://" + hyperlink;
  }

  // Basic URL validation
  const urlPattern = /^(https?:\/\/)[\w.-]+(\.[a-z]{2,})(:[0-9]{1,5})?(\/.*)?$/i;

  if (!urlPattern.test(hyperlink)) {
    alert("Please enter a valid URL (example: https://practicearch.com)");
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
      }
      if (event) event.completed && event.completed();
    }
  );
};
