const defaultNodeWidth = 180;
const nodeHeight = 65;

// Fonction pour estimer la largeur de carte nécessaire selon le texte
function calculateNodeWidth(name) {
  const charCount = name ? name.length : 0;
  // Base minimale de 180px + marge proportionnelle si le nom est long
  return Math.max(defaultNodeWidth, charCount * 9.5 + 24);
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
        // Détermination du genre selon les règles Sosa (Sosa 1 est un homme)
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

    // Construction récursive de l'arbre
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
    
    // Disposition de l'arbre avec espacement dynamique
    const treeLayout = d3.tree().nodeSize([220, nodeHeight + 85]);
    treeLayout(root);

    // Initialisation SVG
    const svg = d3.select("#tree-container").append("svg").attr("width", "100%").attr("height", "100%");
    const g = svg.append("g");

    // Zoom & Pan
    const zoom = d3.zoom().scaleExtent([0.1, 2.5]).on("zoom", (event) => g.attr("transform", event.transform));
    svg.call(zoom);

    // Centrage initial
    const initialX = window.innerWidth / 2;
    const initialY = window.innerHeight - 150;
    svg.call(zoom.transform, d3.zoomIdentity.translate(initialX, initialY).scale(0.8));

    // Liens (Traits noirs épais)
    g.selectAll(".link")
      .data(root.links())
      .enter()
      .append("path")
      .attr("class", "link")
      .attr("d", d3.linkVertical().x(d => d.x).y(d => -d.y));

    // Nœuds (Cartes personnes)
    const nodes = g.selectAll(".node")
      .data(root.descendants())
      .enter()
      .append("g")
      .attr("class", d => `node ${d.data.isMale ? 'node-male' : 'node-female'}`)
      .attr("transform", d => {
        const w = calculateNodeWidth(d.data.personne);
        return `translate(${d.x - w / 2}, ${-d.y - nodeHeight / 2})`;
      });

    // Rectangle avec largeur adaptée au contenu
    nodes.append("rect")
      .attr("width", d => calculateNodeWidth(d.data.personne))
      .attr("height", nodeHeight);

    // SOSA
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

    // Tooltip
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
