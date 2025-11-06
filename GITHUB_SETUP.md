# GitHub Setup Instructions

## Step 1: Create a GitHub Repository

1. Go to https://github.com/new
2. Repository name: `student-risk-analyzer` (or your preferred name)
3. Description: "Student Risk Analyzer - Brukd Consultancy branded web application for analyzing student risk levels"
4. Choose **Public** or **Private**
5. **DO NOT** initialize with README, .gitignore, or license (we already have these)
6. Click **Create repository**

## Step 2: Add Remote and Push

After creating the repository, GitHub will show you commands. Use these commands:

```bash
# Add the remote (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/student-risk-analyzer.git

# Rename branch to main (if needed)
git branch -M main

# Push to GitHub
git push -u origin main
```

## Alternative: Using SSH

If you prefer SSH:

```bash
git remote add origin git@github.com:YOUR_USERNAME/student-risk-analyzer.git
git branch -M main
git push -u origin main
```

## Step 3: Verify

After pushing, visit your repository on GitHub to verify all files are uploaded.

## Optional: Enable GitHub Pages

If you want to host the frontend on GitHub Pages:

1. Go to your repository settings
2. Navigate to **Pages** section
3. Select **main** branch and **/ (root)** folder
4. Click **Save**
5. Your site will be available at: `https://YOUR_USERNAME.github.io/student-risk-analyzer`

Note: The frontend will work, but the API endpoints will need to be hosted separately or use a backend service.

