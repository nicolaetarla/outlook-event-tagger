const DEFAULT_CATEGORY = "MyDefaultTag"; // Pre-defined tag

Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    // No UI initialization needed for commands.js
  }
});

function onMessageSend(event) {
  Office.context.roamingSettings.getAsync('selectedCategory', function(asyncResult) {
    let categoryToApply = DEFAULT_CATEGORY;

    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value) {
      categoryToApply = asyncResult.value;
    }

    // Ensure the category exists in master categories before applying
    Office.context.mailbox.masterCategories.getAsync(function(categoriesResult) {
      if (categoriesResult.status === Office.AsyncResultStatus.Succeeded) {
        const masterCategories = categoriesResult.value.map(cat => cat.displayName);
        if (!masterCategories.includes(categoryToApply)) {
          // If the category doesn't exist, add it to master categories
          Office.context.mailbox.masterCategories.addAsync([categoryToApply], function(addResult) {
            if (addResult.status === Office.AsyncResultStatus.Succeeded) {
              applyCategoryToItem(categoryToApply, event);
            } else {
              console.error("Failed to add category: " + addResult.error.message);
              event.completed({ allowEvent: true }); // Allow send even if category fails
            }
          });
        } else {
          applyCategoryToItem(categoryToApply, event);
        }
      } else {
        console.error("Failed to get master categories: " + categoriesResult.error.message);
        event.completed({ allowEvent: true }); // Allow send even if category fails
      }
    });
  });
}

function applyCategoryToItem(category, event) {
  Office.context.mailbox.item.categories.addAsync([category], function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Category '" + category + "' applied successfully.");
    } else {
      console.error("Failed to apply category: " + asyncResult.error.message);
    }
    event.completed({ allowEvent: true }); // Always allow the event to be sent
  });
}

// This is required by the Office.js runtime for event-based activation
// It makes the onMessageSend function globally accessible.
// @ts-ignore
Office.actions.associate("onMessageSend", onMessageSend);