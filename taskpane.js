// Initialize the Office add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('Outlook add-in loaded successfully');
  }
});

// Function to insert logo with hyperlink
function insertLogo() {
  // Get the current compose item
  Office.context.mailbox.item.body.getAsync(
    Office.CoercionType.Html,
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Prompt user for hyperlink URL
        const hyperlink = prompt("Enter the URL for the hyperlink (e.g., https://practicela.com):");
        
        if (hyperlink) {
          // Validate URL (basic validation)
          let validatedUrl = hyperlink;
          if (!hyperlink.startsWith('http://') && !hyperlink.startsWith('https://')) {
            validatedUrl = 'https://' + hyperlink;
          }
          
          // Create HTML for logo with hyperlink
          const logoHtml = `
            <div style="margin: 10px 0;">
              <a href="${validatedUrl}" target="_blank" style="text-decoration: none;">
                <img src="https://newchigrboy.github.io/practice-outlook-addin/logo.png" 
                     alt="Practice Logo" 
                     style="max-width: 200px; height: auto; border: none;" />
              </a>
            </div>
          `;
          
          // Insert the logo at the current cursor position
          Office.context.mailbox.item.body.setSelectedDataAsync(
            logoHtml,
            { coercionType: Office.CoercionType.Html },
            function (result) {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                console.log('Logo inserted successfully');
                // Optionally show a success message
                alert('Logo inserted successfully!');
              } else {
                console.error('Error inserting logo:', result.error.message);
                alert('Error inserting logo: ' + result.error.message);
              }
            }
          );
        } else {
          alert('No URL provided. Logo insertion cancelled.');
        }
      } else {
        console.error('Error accessing email body:', result.error.message);
        alert('Error accessing email body: ' + result.error.message);
      }
    }
  );
}

// Alternative function if you want to insert at a specific position
function insertLogoAtCursor() {
  const hyperlink = prompt("Enter the URL for the hyperlink:");
  
  if (hyperlink) {
    let validatedUrl = hyperlink;
    if (!hyperlink.startsWith('http://') && !hyperlink.startsWith('https://')) {
      validatedUrl = 'https://' + hyperlink;
    }
    
    const logoHtml = `
      <div>
        <a href="${validatedUrl}" target="_blank">
          <img src="https://newchigrboy.github.io/practice-outlook-addin/logo.png" 
               alt="Practice Architecture Logo" 
               style="max-width: 200px; height: auto;" />
        </a>
      </div>
    `;
    
    Office.context.mailbox.item.body.setSelectedDataAsync(
      logoHtml,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          alert('Logo with hyperlink inserted!');
        } else {
          alert('Error: ' + result.error.message);
        }
      }
    );
  }
}

