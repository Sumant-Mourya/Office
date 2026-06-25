```javascript
rules_version = '2';

service cloud.firestore {
  match /databases/{database}/documents {
    
    // Match any document in the 'users' collection
    match /users/{userId} {
      // Allow the user to read and write their own base document
      allow read, write: if request.auth != null && request.auth.uid == userId;
      
      // Match any documents in subcollections (like the 'pcs' collection) inside the user's document
      match /{document=**} {
        allow read, write: if request.auth != null && request.auth.uid == userId;
      }
    }
    
    // --- Match the app_data collection ---
    match /app_data/{document} {
      // Allow public read access so the website can fetch the pricing and ratings 
      // without needing to be logged in
      allow read: if true;
      
      // Temporarily set to true to allow the website to auto-create missing documents
      // Once created, change this back to: if false;
      allow write: if true;
    }

    // --- Match the top-level reviews collection ---
    match /reviews/{document} {
      allow read: if true;
      
      // Temporarily allow writes so default 6 reviews can auto-create, and users can submit new ones
      allow write: if true; 
    }
    
  }
}
```
