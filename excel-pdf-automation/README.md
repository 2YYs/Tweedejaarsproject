# CAO Analyse & Excel Generator

This project is a web application that analyzes Dutch Collective Labor Agreement (CAO) documents in PDF format, extracts salary increases using an AI model, and generates a summary Excel file.

## How to Run the Application

To run this application, you need to run two separate processes simultaneously in two different terminals: the **Backend API** and the **Frontend Web App**.

### Prerequisites

1.  **Node.js & npm**: Make sure you have Node.js and npm installed. This is for the frontend.
2.  **Python & pip**: Make sure you have Python 3 and pip installed. This is for the backend.
3.  **System Packages**: The Python backend requires two system packages for PDF and OCR processing.
    *   **On macOS**: `brew install tesseract poppler`
    *   **On Linux (Debian/Ubuntu)**: `sudo apt-get install tesseract-ocr poppler-utils`
    *   **Tesseract Language Data**: Ensure you also have the Dutch language pack for Tesseract: `tesseract-nld`. This might be installed with `tesseract` or may require a separate package like `tesseract-ocr-nld`.

---

### Step 1: Run the Backend API (Terminal 1)

The backend is a Python server that handles the heavy lifting of PDF analysis.

1.  **Navigate to the backend directory:**
    ```bash
    cd src/backend
    ```

2.  **(First time only) Install Python dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the server:**
    ```bash
    python server.py
    ```

4.  **Keep this terminal open!** The server will start logging information. If it starts correctly, you will see lines like `* Running on http://127.0.0.1:5001`. The frontend needs this server to be running to function. All detailed processing logs will appear here.

---

### Step 2: Run the Frontend App (Terminal 2)

The frontend is a React application that provides the user interface in your browser.

1.  **Navigate to the project's root directory** (if you're not already there).

2.  **(First time only) Install Node.js dependencies:**
    ```bash
    npm install
    ```

3.  **Run the frontend:**
    ```bash
    npm run dev
    ```

4.  The terminal will give you a local URL to open in your browser, usually `http://localhost:5173`. Open this URL.

You should now be able to use the application by uploading your PDF files. Please check the output in **Terminal 1** for the detailed backend logs, as this will show exactly what the application is doing.

## Project info

**URL**: https://lovable.dev/projects/bb3b8d9d-a8d2-44fb-9e61-f1b5091af101

## How can I edit this code?

There are several ways of editing your application.

**Use Lovable**

Simply visit the [Lovable Project](https://lovable.dev/projects/bb3b8d9d-a8d2-44fb-9e61-f1b5091af101) and start prompting.

Changes made via Lovable will be committed automatically to this repo.

**Use your preferred IDE**

If you want to work locally using your own IDE, you can clone this repo and push changes. Pushed changes will also be reflected in Lovable.

The only requirement is having Node.js & npm installed - [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating)

Follow these steps:

```sh
# Step 1: Clone the repository using the project's Git URL.
git clone <YOUR_GIT_URL>

# Step 2: Navigate to the project directory.
cd <YOUR_PROJECT_NAME>

# Step 3: Install the necessary dependencies.
npm i

# Step 4: Start the development server with auto-reloading and an instant preview.
npm run dev
```

**Edit a file directly in GitHub**

- Navigate to the desired file(s).
- Click the "Edit" button (pencil icon) at the top right of the file view.
- Make your changes and commit the changes.

**Use GitHub Codespaces**

- Navigate to the main page of your repository.
- Click on the "Code" button (green button) near the top right.
- Select the "Codespaces" tab.
- Click on "New codespace" to launch a new Codespace environment.
- Edit files directly within the Codespace and commit and push your changes once you're done.

## What technologies are used for this project?

This project is built with:

- Vite
- TypeScript
- React
- shadcn-ui
- Tailwind CSS

## How can I deploy this project?

Simply open [Lovable](https://lovable.dev/projects/bb3b8d9d-a8d2-44fb-9e61-f1b5091af101) and click on Share -> Publish.

## Can I connect a custom domain to my Lovable project?

Yes, you can!

To connect a domain, navigate to Project > Settings > Domains and click Connect Domain.

Read more here: [Setting up a custom domain](https://docs.lovable.dev/tips-tricks/custom-domain#step-by-step-guide)
