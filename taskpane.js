Office.onReady(() => {
  // Office is ready, functions can now be called
  console.log("Practice add-in ready");
});

function insertLogo(event) {
  // Fixed image URL (GitHub Pages hosted)
  const imageUrl = "https://newchigrboy.github.io/practice-outlook-addin/logo.png";

  // Ask for hyperlink
  const hyperlink = prompt("Enter the hyperlink (e.g., https://www.practicearch.com):");
  if (!hyperlink) {
    if (event) event.completed();
    return;
  }

  // Insert the image wrapped in a link
  Office.context.mailbox.item.body.setSelectedDataAsync(
    `<a href="${hyperlink}" target="_blank"><img src="${imageUrl}" style="max-width:200px;"/></a>`,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to insert logo: " + asyncResult.error.message);
      }
      if (event) event.completed(); // important: tell Office the command is done
    }
  );
}

// Expose function so it's callable by the manifest Action
if (typeof window !== "undefined") {
  window.insertLogo = insertLogo;
}

