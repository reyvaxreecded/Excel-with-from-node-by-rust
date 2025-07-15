const path = require("path");
const fs = require("fs");

// Charger le module Rust
let rustCalamine; // Changé pour correspondre à la variable utilisée plus loin

try {
  // Chemin vers la bibliothèque compilée
  const rustAddonDir = path.join(__dirname, "../rust-addon");
  const npmDir = path.join(rustAddonDir, "npm");

  // Créer le répertoire npm s'il n'existe pas
  if (!fs.existsSync(npmDir)) {
    fs.mkdirSync(npmDir, { recursive: true });
  }

  // Déterminer le nom du fichier de la bibliothèque compilée
  let libExt;
  if (process.platform === "win32") {
    libExt = ".dll";
  } else if (process.platform === "darwin") {
    libExt = ".dylib";
  } else {
    libExt = ".so";
  }

  // Chemin source de la bibliothèque compilée
  const libSrcPath = path.join(
    rustAddonDir,
    "target",
    "release",
    `libexcel_manager${libExt}`
  );

  // Nom du fichier cible de la bibliothèque
  let targetFileName;
  if (process.platform === "win32") {
    targetFileName = "excel_manager.win32-x64-msvc.node";
  } else if (process.platform === "darwin") {
    if (process.arch === "arm64") {
      targetFileName = "excel_manager.darwin-arm64.node";
    } else {
      targetFileName = "excel_manager.darwin-x64.node";
    }
  } else {
    // Pour Linux
    if (process.arch === "arm64") {
      targetFileName = "excel_manager.linux-arm64-gnu.node";
    } else {
      targetFileName = "excel_manager.linux-x64-gnu.node";
    }
  }

  const libTargetPath = path.join(npmDir, targetFileName);

  // Copier la bibliothèque compilée si nécessaire
  if (
    fs.existsSync(libSrcPath) &&
    (!fs.existsSync(libTargetPath) ||
      fs.statSync(libSrcPath).mtime > fs.statSync(libTargetPath).mtime)
  ) {
    fs.copyFileSync(libSrcPath, libTargetPath);
    console.log(`Bibliothèque copiée vers ${libTargetPath}`);
  }

  // Charger le module
  rustCalamine = require("../rust-addon");
  console.log("Module Rust chargé avec succès");

  // Afficher les fonctions disponibles pour le débogage
  console.log("Fonctions disponibles:", Object.keys(rustCalamine));

  // Vérifier que le fichier Excel existe
  const excelFilePath = path.join(
    rustAddonDir,
    "src",
    "assets",
    "excel_files",
    "Cotation-Parc-Batterie-01072025.xlsx"
  );
  if (!fs.existsSync(excelFilePath)) {
    console.error(
      `Le fichier Excel n'existe pas à l'emplacement: ${excelFilePath}`
    );
    console.log(
      "Veuillez créer le dossier et y placer un fichier Excel de test."
    );
    // Créer le dossier s'il n'existe pas
    const excelDir = path.join(rustAddonDir, "src", "assets", "excel_files");
    if (!fs.existsSync(excelDir)) {
      fs.mkdirSync(excelDir, { recursive: true });
      console.log(`Dossier créé: ${excelDir}`);
    }
    return;
  }

  try {
    const data = rustCalamine.read_excel_file("Cotation-Parc-Batterie-01072025.xlsx", "Feuil1");
    console.log("Data :", data); // Affiche la première ligne de données
  } catch (error) {
    console.error("Erreur lors de la lecture du fichier Excel :", error);
  }
} catch (error) {
  console.error("Erreur lors du chargement du module Rust :", error);
}
