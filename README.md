# 🎟️ E-Ticket Pro — Solution Globale de Billetterie, Contrôle d'Accès & Cashless

> **Normes FIFA Ready | Appel d'Offres N° 02/2026/GSC — Grand Stade de Casablanca**  
> *Développé par Wiam El Arroussi — SOMAYAR S.A.R.L. AU*

---

## 🚀 Comment Lancer et Tester l'Application ?

Votre encadrant ou évaluateur dispose de **3 méthodes simples** pour exécuter l'application sur sa machine :

---

### 🟢 Méthode 1 : Lancement Immédiat Sans Installation (Recommandé)

Si vous voulez simplement ouvrir l'application dans votre navigateur sans rien installer :

1. **Téléchargez le projet** (ou clonez le dépôt) :
   ```bash
   git clone https://github.com/wiam-elarroussi/billeterie.git
   cd billeterie
   ```
2. **Ouvrez simplement le fichier `index.html`** dans votre navigateur (Double-cliquez sur `index.html` ou glissez-le dans Chrome / Edge / Firefox).
   *Ou lancez un serveur local rapide avec Node.js :*
   ```bash
   node server.js
   ```
3. Rendez-vous sur **`http://localhost:8099/`** dans votre navigateur.

---

### 🐳 Méthode 2 : Lancement Industriel Complet (Via Docker Compose)

Si vous avez **Docker Desktop** installé sur votre machine, cette commande déploie l'intégralité de la stack (PostgreSQL 16 + Redis + Backend NestJS/gRPC + Frontend) en un seul clic :

```bash
# 1. Cloner le dépôt
git clone https://github.com/wiam-elarroussi/billeterie.git
cd billeterie

# 2. Lancer toute l'infrastructure avec Docker
docker-compose up -d --build
```

**Accès aux services :**
- 🌐 **Frontend Web App** : [http://localhost:8099/](http://localhost:8099/)
- 📡 **Backend REST API** : `http://localhost:4000/api/v1`
- ⚡ **Service gRPC Gate Control** : `localhost:50051`

---

### 💻 Méthode 3 : Lancement Manuel des Microservices (Node.js)

Si vous préférez exécuter les composants séparément sans Docker :

#### 1. Lancer le Frontend Web :
```bash
node server.js
```
*Accès : [http://localhost:8099/](http://localhost:8099/)*

#### 2. Lancer le Backend NestJS & gRPC :
```bash
cd backend
npm install
npx prisma generate
npm run start:dev
```
*Accès REST API : [http://localhost:4000/api/v1](http://localhost:4000/api/v1)*

---

## 🛠️ Stack Technique de Production

- **Frontend** : Next.js 14, TailwindCSS, HTML5/JS, Chart.js, QRCode.js (100% Responsive & PWA Ready).
- **Backend Microservices** : NestJS (TypeScript), Go (Golang), REST API HTTP v1 & Microservice gRPC.
- **Bases de Données & Cache** : PostgreSQL 16 3NF + Extension PostGIS, Redis Cluster (Lock sièges < 2ms).
- **Edge Computing Stade** : Serveur physique Dell PowerEdge R360 (Validation < 0.15s en mode Offline).
- **Sécurité Anti-Fraude** : Algorithme Checksum Mode 4 (`ETK-2026-XXXX`).

---

## 📄 Documentation Technique & Fichiers Word Inclus

L'intégralité des dossiers d'ingénierie se trouve dans le projet sous forme de fichiers Word `.docx` :
1. 📘 **`Rapport_Executive_Direction_Application_Complete_E-Ticket_Pro.docx`** — Synthèse pour la Direction Générale.
2. 📘 **`Dossier_Conception_Technique_Stack_Data_API.docx`** — Stack, Schéma SQL ERD et Spécifications des APIs REST & gRPC.
3. 📘 **`Guide_Utilisation_Application_Web_Complete_E-Ticket_Pro.docx`** — Manuel Utilisateur Officiel.
4. 📘 **`Cahier_des_Charges_E-Ticket_Pro.docx`** — Spécifications Fonctionnelles & Techniques.
5. 📘 **`Definition_Besoins_Utilisateurs_Cibles.docx`** — Cartographie des Personas.
6. 📘 **`Fonctionnalites_MVP_E-Ticket_Pro.docx`** — Périmètre du MVP.
