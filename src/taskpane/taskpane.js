Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    const saveButton = document.getElementById('saveButton');
    const categorySelect = document.getElementById('categorySelect');

    saveButton.onclick = saveCategorySetting;
    saveButton.disabled = true; // Disable initially

    categorySelect.onchange = function() {
      saveButton.disabled = (categorySelect.value === "");
    };

    loadCategories();
  }
});

async function loadCategories() {
  try {
    // Get all categories from the mailbox
    const categories = await Office.context.mailbox.masterCategories.getAsync();
    const categorySelect = document.getElementById('categorySelect');

    if (categories.status === Office.AsyncResultStatus.Succeeded) {
      categories.value.forEach(category => {
        const option = document.createElement('option');
        option.value = category.displayName;
        option.textContent = category.displayName;
        categorySelect.appendChild(option);
      });

      // Load previously saved category
      Office.context.roamingSettings.getAsync('selectedCategory', function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
          categorySelect.value = asyncResult.value;
          document.getElementById('saveButton').disabled = false; // Enable if a category is already selected
        } else {
          document.getElementById('saveButton').disabled = true; // Keep disabled if no category is selected
        }
      });
    } else {
      console.error("Failed to get master categories: " + categories.error.message);
      displayMessage("Error loading categories.", true);
    }
  } catch (error) {
    console.error("Error in loadCategories: " + error);
    displayMessage("Error loading categories.", true);
  }
}

function saveCategorySetting() {
  const categorySelect = document.getElementById('categorySelect');
  const selectedCategory = categorySelect.value;

  Office.context.roamingSettings.setAsync('selectedCategory', selectedCategory, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      displayMessage('Category setting saved successfully!');
    } else {
      console.error("Failed to save setting: " + asyncResult.error.message);
      displayMessage('Error saving setting.', true);
    }
  });
}

function displayMessage(message, isError = false) {
  const messageDiv = document.getElementById('message');
  messageDiv.textContent = message;
  messageDiv.style.color = isError ? 'red' : 'green';
}