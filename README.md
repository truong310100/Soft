…or create a new repository on the command line
echo "# Soft" >> README.md
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/truong310100/Soft.git
git push -u origin main

…or push an existing repository from the command line
git remote add origin https://github.com/truong310100/Soft.git
git branch -M main
git push -u origin main

git add .
git commit -m "update"
git push -u origin main
