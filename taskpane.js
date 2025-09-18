Office.onReady(() => {
  console.log("Practice add-in ready");
});

function insertLogoFromInput(event) {
  const hyperlink = document.getElementById("linkInput").value.trim();
  if (!hyperlink) {
    alert("Please enter a hyperlink before inserting the logo.");
    if (event) event.completed();
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
      }
      if (event) event.completed();
    }
  );
}
