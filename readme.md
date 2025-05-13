# Gestionnaire de fichiers Excel en Rust avec intégrations Node.js et NestJS

Ce projet est une bibliothèque Rust qui permet de manipuler et lire des fichiers Excel de manière efficace. Il utilise la bibliothèque `calamine` en Rust et fournit des bindings pour Node.js et NestJS.

## Structure du projet

```
.
├── nestJs/               # Implémentation NestJS
├── node/                 # Implémentation Node.js pure
└── rust-addon/          # Code Rust principal
```

## Fonctionnalités

- Lecture de fichiers Excel (.xlsx, .xls)
- Manipulation des données de feuilles de calcul
- Intégration native avec Node.js via N-API
- Support complet pour NestJS

## Prérequis

- Rust (édition 2021 ou supérieure)
- Node.js (v14 ou supérieure)
- npm ou yarn
- Cargo (gestionnaire de paquets Rust)

## Installation

### Pour l'utilisation avec Node.js

```bash
cd node
npm install
```

### Pour l'utilisation avec NestJS

```bash
cd nestJs/nest-app
npm install
```

## Utilisation

### Version Node.js

```javascript
const excelManager = require('./rust-addon');

// Exemple d'utilisation
const data = excelManager.readExcelFile('chemin/vers/fichier.xlsx');
```

### Version NestJS

```typescript
import { ExcelService } from './excel.service';

@Controller()
export class AppController {
  constructor(private readonly excelService: ExcelService) {}

  @Get('read-excel')
  readExcel() {
    return this.excelService.readExcelFile('chemin/vers/fichier.xlsx');
  }
}
```

## Développement

Pour compiler le code Rust :

```bash
cd rust-addon
cargo build --release
```

## Tests

Pour exécuter les tests :

```bash
# Tests Rust
cd rust-addon
cargo test

# Tests Node.js
cd node
npm test

# Tests NestJS
cd nestJs/nest-app
npm test
```

## Licence

MIT

## Contribution

Les contributions sont les bienvenues ! N'hésitez pas à ouvrir une issue ou une pull request.