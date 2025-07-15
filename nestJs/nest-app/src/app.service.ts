import { Injectable } from '@nestjs/common';
import type {
  read_excel_file,
  UpdateCell,
  upsert_row,
} from 'excel-manager-native';
import * as fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import * as path from 'path';

const _require = createRequire(__filename);

interface ExcelManagerModule {
  read_excel_file: typeof read_excel_file;
  upsert_row: typeof upsert_row;
}
@Injectable()
export class AppService {
  private excelManager: ExcelManagerModule | null = null;
  private moduleLoadAttempted = false;
  private baseExcelPath: string = '';

  async onModuleInit() {
    await this.initializeRustModule();
  }

  onModuleDestroy() {
    this.excelManager = null;
  }

  private async initializeRustModule() {
    try {
      // Chemin du module Rust
      const rustAddonDir = join(process.cwd(), '../../rust-addon');
      const npmDir = join(rustAddonDir, 'npm');

      // Créer le répertoire npm s'il n'existe pas
      try {
        await fs.mkdir(npmDir, { recursive: true });
      } catch (err) {
        // Ignorer l'erreur si le dossier existe déjà
        console.error(err);
      }

      // Déterminer l'extension en fonction du système d'exploitation
      let libExt;
      if (process.platform === 'win32') {
        libExt = '.dll';
      } else if (process.platform === 'darwin') {
        libExt = '.dylib';
      } else {
        libExt = '.so';
      }

      // Chemin source et cible
      const libSrcPath = join(
        rustAddonDir,
        'target',
        'release',
        `libexcel_manager${libExt}`,
      );

      // Nom du fichier cible
      let targetFileName;
      if (process.platform === 'win32') {
        targetFileName = 'excel_manager.win32-x64-msvc.node';
      } else if (process.platform === 'darwin') {
        if (process.arch === 'arm64') {
          targetFileName = 'excel_manager.darwin-arm64.node';
        } else {
          targetFileName = 'excel_manager.darwin-x64.node';
        }
      } else {
        // Pour Linux
        if (process.arch === 'arm64') {
          targetFileName = 'excel_manager.linux-arm64-gnu.node';
        } else {
          targetFileName = 'excel_manager.linux-x64-gnu.node';
        }
      }

      const libTargetPath = join(npmDir, targetFileName);

      // Vérifier si la bibliothèque source existe
      let sourceExists = false;
      try {
        await fs.stat(libSrcPath);
        sourceExists = true;
      } catch (error) {
        console.error(error);
        sourceExists = false;
      }

      if (!sourceExists) {
        throw new Error(`La bibliothèque source n'existe pas: ${libSrcPath}`);
      }

      // Vérifier si la bibliothèque cible existe
      let targetExists = false;
      try {
        await fs.stat(libTargetPath);
        targetExists = true;
      } catch (error) {
        console.error(error);
        targetExists = false;
      }

      // Copier la bibliothèque si nécessaire
      if (
        !targetExists ||
        (sourceExists &&
          (await fs.stat(libSrcPath)).mtime >
            (await fs.stat(libTargetPath)).mtime)
      ) {
        await fs.copyFile(libSrcPath, libTargetPath);
        console.log(`Bibliothèque copiée vers ${libTargetPath}`);
      }

      // Utiliser createRequire pour charger le module natif
      try {
        const nativeModule = _require(libTargetPath);
        this.excelManager = nativeModule;
        this.moduleLoadAttempted = true;

        console.log('Module Excel Rust chargé avec succès');
        console.log('Fonctions disponibles:', Object.keys(this.excelManager));
      } catch (error) {
        console.error('Erreur lors du chargement du module natif:', error);
        throw error;
      }

      // Définir le chemin de base pour les fichiers Excel
      this.baseExcelPath = join(rustAddonDir, 'src', 'assets', 'excel_files');
      console.log('Base path for Excel files:', this.baseExcelPath);
      try {
        await fs.mkdir(this.baseExcelPath, { recursive: true });
      } catch (err) {
        console.error('Erreur lors de la création du répertoire Excel:', err);
        // Ignorer l'erreur si le dossier existe déjà
      }
    } catch (error) {
      console.error("Erreur lors de l'initialisation du module Rust:", error);
      this.moduleLoadAttempted = true;

      if (process.env.NODE_ENV === 'development') {
        console.warn(
          'Activation du mode fallback JavaScript pour le développement',
        );
        this.initJsFallback();
      } else {
        throw error;
      }
    }
  }

  private async loadNativeModule(
    modulePath: string,
  ): Promise<ExcelManagerModule> {
    try {
      // Utiliser import() dynamique au lieu de require()
      // Nous devons utiliser eval ici car TypeScript ne peut pas vérifier le chemin dynamique
      // eslint-disable-next-line @typescript-eslint/no-implied-eval
      const importFunction = new Function(
        'modulePath',
        'return import(modulePath)',
      );
      const module = await importFunction(modulePath);
      return module as ExcelManagerModule;
    } catch (error) {
      console.error('Erreur lors du chargement du module natif:', error);
      throw error;
    }
  }

  private initJsFallback() {
    // Implémentation JS basique qui sera utilisée si le module natif n'est pas disponible
    this.excelManager = {
      read_excel_file: (fileName: string): string[][] => {
        console.log(`[JS FALLBACK] Lecture de ${fileName}`);
        return [
          ['Données', 'de', 'secours'],
          ['pour', 'le', 'développement'],
        ];
      },
      upsert_row: (
        fileName: string,
        sheetName: string,
        row: UpdateCell[],
      ): void => {
        console.log(
          `[JS FALLBACK] Mise à jour du fichier ${fileName}, feuille ${sheetName}`,
        );
        console.log('Données:', row);
      },
    };
  }

  /**
   * Vérifie si un fichier Excel existe
   */
  async excelFileExists(fileName: string): Promise<boolean> {
    const filePath = path.join(this.baseExcelPath, fileName);
    console.log("Vérification de l'existence du fichier:", filePath);
    try {
      await fs.access(filePath);
      return true;
    } catch (error) {
      console.error(error);
      return false;
    }
  }

  /**
   * Lit le contenu d'un fichier Excel
   */
  async readExcelFile(fileName: string): Promise<string[][]> {
    if (!this.excelManager && this.moduleLoadAttempted) {
      await this.initializeRustModule();
    }

    if (!this.excelManager) {
      throw new Error("Le module Excel n'est pas initialisé");
    }

    if (!(await this.excelFileExists(fileName))) {
      throw new Error(`Le fichier Excel "${fileName}" n'existe pas`);
    }

    try {
      return this.excelManager.read_excel_file(fileName, 'Feuil1');
    } catch (error) {
      console.error('Erreur lors de la lecture du fichier Excel:', error);
      throw error;
    }
  }

  /**
   * Met à jour ou insère une ligne dans un fichier Excel
   */
  async upsertRow(
    fileName: string,
    sheetName: string,
    row: UpdateCell[],
  ): Promise<void> {
    if (!this.excelManager && this.moduleLoadAttempted) {
      await this.initializeRustModule();
    }

    if (!this.excelManager) {
      throw new Error("Le module Excel n'est pas initialisé");
    }

    if (!(await this.excelFileExists(fileName))) {
      throw new Error(`Le fichier Excel "${fileName}" n'existe pas`);
    }

    try {
      return this.excelManager.upsert_row(fileName, sheetName, row);
    } catch (error) {
      console.error('Erreur lors de la mise à jour du fichier Excel:', error);
      throw error;
    }
  }
}

export { UpdateCell };
