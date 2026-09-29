const defaultNodeWidth = 180;
const nodeHeight = 65;

// Fonction pour calculer la largeur de la carte selon le texte
function calculateNodeWidth(name) {
  const charCount = name ? String(name).length : 0;
  // Largeur dynamique avec un minimum de 180px
  return Math.max(defaultNodeWidth, charCount * 9 + 30);
}

// Charger le fichier Excel 'data/arbre.xlsx'
fetch('data/arbre.xlsx')
  .then(res => res.arrayBuffer())
  .then(buffer => {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    const dataBySosa = {};

    rawData.forEach(row => {
      const sosa = parseInt(row['N° SOSA'], 10);
      
      if (!isNaN(sosa)) {
        // Détermination du genre selon la règle Sosa (Sosa 1 est un homme)
        const isMale = sosa === 1 || sosa % 2 === 0;

        dataBySosa[sosa] = {
          sosa: sosa,
          generation: row['Rang / Gén.'],
          personne: String(row['Personne']).trim(),
          isMale: isMale,
          dateNaissance: row['Date de naissance'],
          lieuNaissance: row['Lieu de naissance'],
          conjoint: row['Conjoint(e)'],
          dateMariage: row['Date de mariage'],
          lieuMariage: row['Lieu de mariage'],
          nbEnfants: row['Nb enfants'],
          dateDeces: row['Date de décès'],
          lieuDeces: row['Lieu de décès'],
          age: row['Âge'],
          profession: row['Profession'],
          notes: row['Notes'],
          fatherSosa: sosa * 2,
          motherSosa: (sosa * 2) + 1
        };
      }
    });

    // Construction récursive de la hiérarchie
    function buildHierarchy(sosa) {
      const person = dataBySosa[sosa];
      if (!person) return null;

      const node = { ...person, children: [] };

      if (dataBySosa[person.fatherSosa]) {
        const father = buildHierarchy(person.fatherSosa);
        if (father) node.children.push(father);
      }

      if (dataBySosa[person.motherSosa]) {
        const mother = buildHierarchy(person.motherSosa);
        if (mother) node.children.push(mother);
      }

      return node;
    }

    const rootData = buildHierarchy(1);
    if (!rootData) {
      console.error("Aucune personne avec le N° SOSA 1 n'a été trouvée.");
      return;
    }

    const root = d3.hierarchy(rootData);
    
    // Configuration du layout avec espacement dynamique entre nœuds voisins
    const treeLayout = d3.tree()
      .nodeSize([280, nodeHeight + 85]) // Augmentation du pas de base X (280px)
      .separation((a, b) => {
        // Calcul du besoin de largeur combinée des deux nœuds adjacents
        const widthA = calculateNodeWidth(a.data.personne);
        const widthB = calculateNodeWidth(b.data.personne);
        const requiredSpacing = (widthA + widthB) / 2 + 40; // 40px d'espace de sécurité entre cartes
        const baseStep = 280;
        
        const factor = requiredSpacing / baseStep;
        // Si a et b sont de mêmes parents, on applique la distance calculée, sinon un peu plus
        return a.parent === b.parent ? Math.max(1.1, factor) : Math.max(1.3, factor * 1.15);
      });

    treeLayout(root);

    // Initialisation du SVG
    const svg = d3.select("#tree-container").append("svg").attr("width", "100%").attr("height", "100%");
    const g = svg.append("g");

    // Gestion du Zoom et Pan
    const zoom = d3.zoom().scaleExtent([0.1, 2.5]).on("zoom", (event) => g.attr("transform", event.transform));
    svg.call(zoom);

    // Centrage initial sur le SOSA 1
    const initialX = window.innerWidth / 2;
    const initialY = window.innerHeight - 150;
    svg.call(zoom.transform, d3.zoomIdentity.translate(initialX, initialY).scale(0.8));

    // Dessin des liens (Traits noirs)
    g.selectAll(".link")
      .data(root.links())
      .enter()
      .append("path")
      .attr("class", "link")
      .attr("d", d3.linkVertical().x(d => d.x).y(d => -d.y));

    // Dessin des cartes (Nœuds)
    const nodes = g.selectAll(".node")
      .data(root.descendants())
      .enter()
      .append("g")
      .attr("class", d => `node ${d.data.isMale ? 'node-male' : 'node-female'}`)
      .attr("transform", d => {
        const w = calculateNodeWidth(d.data.personne);
        return `translate(${d.x - w / 2}, ${-d.y - nodeHeight / 2})`;
      });

    // Rectangle avec largeur adaptée
    nodes.append("rect")
      .attr("width", d => calculateNodeWidth(d.data.personne))
      .attr("height", nodeHeight);

    // N° Sosa
    nodes.append("text")
      .attr("class", "sosa")
      .attr("x", 10)
      .attr("y", 16)
      .text(d => `Sosa ${d.data.sosa}`);

    // Nom complet
    nodes.append("text")
      .attr("class", "name")
      .attr("x", 10)
      .attr("y", 35)
      .text(d => d.data.personne);

    // Dates
    nodes.append("text")
      .attr("class", "dates")
      .attr("x", 10)
      .attr("y", 53)
      .text(d => {
        const n = d.data.dateNaissance ? String(d.data.dateNaissance).split('/').pop() : '?';
        const dcs = d.data.dateDeces ? String(d.data.dateDeces).split('/').pop() : '';
        return `${n} - ${dcs}`;
      });

    // Info-bulle au survol (Tooltip)
    const tooltip = d3.select("#tooltip");

    nodes.on("mouseover", (event, d) => {
      const p = d.data;

      tooltip.html(`
        <div class="tooltip-header">
          <div class="tooltip-title">${p.personne}</div>
          <div class="tooltip-subtitle">Sosa ${p.sosa} ${p.generation ? '• Génération ' + p.generation : ''}</div>
        </div>
        <div class="tooltip-row"><strong>Naissance :</strong> ${p.dateNaissance || '?'} ${p.lieuNaissance ? 'à ' + p.lieuNaissance : ''}</div>
        <div class="tooltip-row"><strong>Décès :</strong> ${p.dateDeces || 'Inconnu'} ${p.lieuDeces ? 'à ' + p.lieuDeces : ''} ${p.age ? '(' + p.age + ' ans)' : ''}</div>
        <div class="tooltip-row"><strong>Profession :</strong> ${p.profession || 'Non renseignée'}</div>
        <div class="tooltip-row"><strong>Union :</strong> ${p.conjoint || 'Non renseignée'} ${p.dateMariage ? '(' + p.dateMariage + ')' : ''} ${p.lieuMariage ? 'à ' + p.lieuMariage : ''}</div>
        ${p.notes ? `<div class="tooltip-notes"><strong>Note :</strong> ${p.notes}</div>` : ''}
      `);

      tooltip.style("opacity", 1);
    })
    .on("mousemove", (event) => {
      tooltip.style("left", (event.pageX + 15) + "px").style("top", (event.pageY - 15) + "px");
    })
    .on("mouseout", () => {
      tooltip.style("opacity", 0);
    });

  })
  .catch(err => console.error("Erreur de chargement du fichier Excel :", err));
