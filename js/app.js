const nodeWidth = 180;
const nodeHeight = 60;

// Charger le fichier Excel 'data/arbre.xlsx'
fetch('data/arbre.xlsx')
  .then(res => res.arrayBuffer())
  .then(buffer => {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    // Conversion de la feuille Excel en tableau d'objets JS
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    // Traitement des données et correspondance exacte avec tes colonnes
    const dataBySosa = {};

    rawData.forEach(row => {
      const sosa = parseInt(row['N° SOSA'], 10);
      
      // On conserve uniquement les personnes possédant un N° SOSA valide
      if (!isNaN(sosa)) {
        dataBySosa[sosa] = {
          sosa: sosa,
          generation: row['Rang / Gén.'],
          personne: row['Personne'],
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
          // Paternité / Maternité théorique Sosa
          fatherSosa: sosa * 2,
          motherSosa: (sosa * 2) + 1
        };
      }
    });

    // Construction récursive de l'arbre binaire
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
    const treeLayout = d3.tree().nodeSize([nodeWidth + 40, nodeHeight + 80]);
    treeLayout(root);

    // Initialisation du canevas SVG
    const svg = d3.select("#tree-container").append("svg").attr("width", "100%").attr("height", "100%");
    const g = svg.append("g");

    // Gestion du Zoom et Pan
    const zoom = d3.zoom().scaleExtent([0.1, 2.5]).on("zoom", (event) => g.attr("transform", event.transform));
    svg.call(zoom);

    // Centrage initial sur le SOSA 1 (en bas au centre)
    const initialX = window.innerWidth / 2;
    const initialY = window.innerHeight - 150;
    svg.call(zoom.transform, d3.zoomIdentity.translate(initialX, initialY).scale(0.8));

    // Tracé des branches (liens)
    g.selectAll(".link")
      .data(root.links())
      .enter()
      .append("path")
      .attr("class", "link")
      .attr("d", d3.linkVertical().x(d => d.x).y(d => -d.y));

    // Tracé des cartes (nœuds)
    const nodes = g.selectAll(".node")
      .data(root.descendants())
      .enter()
      .append("g")
      .attr("class", "node")
      .attr("transform", d => `translate(${d.x - nodeWidth / 2}, ${-d.y - nodeHeight / 2})`);

    nodes.append("rect").attr("width", nodeWidth).attr("height", nodeHeight);

    nodes.append("text").attr("class", "sosa").attr("x", 8).attr("y", 15).text(d => `Sosa ${d.data.sosa}`);
    nodes.append("text").attr("class", "name").attr("x", 8).attr("y", 32).text(d => d.data.personne);
    nodes.append("text").attr("class", "dates").attr("x", 8).attr("y", 48).text(d => {
      const n = d.data.dateNaissance ? String(d.data.dateNaissance).split('/').pop() : '?';
      const dcs = d.data.dateDeces ? String(d.data.dateDeces).split('/').pop() : '';
      return `${n} - ${dcs}`;
    });

    // Gestion du survol (Hover) pour l'Info-bulle
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