const defaultNodeWidth = 180;
const nodeHeight = 65;

// Fonction pour calculer la largeur de la carte selon le texte
function calculateNodeWidth(name) {
  const charCount = name ? String(name).length : 0;
  return Math.max(defaultNodeWidth, charCount * 9 + 30);
}

// Fonction robuste pour extraire l'année ou la date lisible
function extractYear(val) {
  if (!val) return '';
  const s = String(val).trim();
  if (s.includes('/')) {
    const parts = s.split('/');
    return parts[parts.length - 1].trim();
  }
  const match = s.match(/\b(1[2-9]\d{2}|20\d{2})\b/);
  if (match) return match[0];
  return s;
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
    let maxGenFound = 1;

    rawData.forEach(row => {
      const sosa = parseInt(row['N° SOSA'], 10);
      
      if (!isNaN(sosa)) {
        const isMale = sosa === 1 || sosa % 2 === 0;
        const gen = parseInt(row['Rang / Gén.'], 10) || (Math.floor(Math.log2(sosa)) + 1);
        if (gen > maxGenFound) maxGenFound = gen;

        dataBySosa[sosa] = {
          sosa: sosa,
          generation: gen,
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

    const genSelect = d3.select("#gen-select");
    if (genSelect.node() && genSelect.selectAll("option").empty()) {
      for (let i = 1; i <= maxGenFound; i++) {
        genSelect.append("option")
          .attr("value", i)
          .text(`${i} génération${i > 1 ? 's' : ''}`)
          .property("selected", i === maxGenFound);
      }
    }

    const svg = d3.select("#tree-container").append("svg").attr("width", "100%").attr("height", "100%");
    const g = svg.append("g");

    const zoom = d3.zoom().scaleExtent([0.05, 3]).on("zoom", (event) => g.attr("transform", event.transform));
    svg.call(zoom);

    function renderTree(maxGen) {
      g.selectAll("*").remove();

      function buildHierarchy(sosa) {
        const person = dataBySosa[sosa];
        if (!person || person.generation > maxGen) return null;

        const node = { ...person, children: [] };

        if (dataBySosa[person.fatherSosa] && dataBySosa[person.fatherSosa].generation <= maxGen) {
          const father = buildHierarchy(person.fatherSosa);
          if (father) node.children.push(father);
        }

        if (dataBySosa[person.motherSosa] && dataBySosa[person.motherSosa].generation <= maxGen) {
          const mother = buildHierarchy(person.motherSosa);
          if (mother) node.children.push(mother);
        }

        return node;
      }

      const rootData = buildHierarchy(1);
      if (!rootData) return;

      const root = d3.hierarchy(rootData);
      
      const treeLayout = d3.tree()
        .nodeSize([280, nodeHeight + 85])
        .separation((a, b) => {
          const widthA = calculateNodeWidth(a.data.personne);
          const widthB = calculateNodeWidth(b.data.personne);
          const requiredSpacing = (widthA + widthB) / 2 + 40;
          const factor = requiredSpacing / 280;
          return a.parent === b.parent ? Math.max(1.1, factor) : Math.max(1.3, factor * 1.15);
        });

      treeLayout(root);

      const descendants = root.descendants();

      g.selectAll(".link")
        .data(root.links())
        .enter()
        .append("path")
        .attr("class", "link")
        .attr("d", d3.linkVertical().x(d => d.x).y(d => -d.y));

      const nodes = g.selectAll(".node")
        .data(descendants)
        .enter()
        .append("g")
        .attr("class", d => `node ${d.data.isMale ? 'node-male' : 'node-female'}`)
        .attr("transform", d => {
          const w = calculateNodeWidth(d.data.personne);
          return `translate(${d.x - w / 2}, ${-d.y - nodeHeight / 2})`;
        });

      nodes.append("rect")
        .attr("width", d => calculateNodeWidth(d.data.personne))
        .attr("height", nodeHeight);

      nodes.append("text").attr("class", "sosa").attr("x", 10).attr("y", 16).text(d => `Sosa ${d.data.sosa}`);
      nodes.append("text").attr("class", "name").attr("x", 10).attr("y", 35).text(d => d.data.personne);
      
      nodes.append("text").attr("class", "dates").attr("x", 10).attr("y", 53).text(d => {
        const n = extractYear(d.data.dateNaissance) || '?';
        const dcs = extractYear(d.data.dateDeces);
        return `${n} - ${dcs}`;
      });

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

      let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
      descendants.forEach(d => {
        const w = calculateNodeWidth(d.data.personne);
        if (d.x - w / 2 < minX) minX = d.x - w / 2;
        if (d.x + w / 2 > maxX) maxX = d.x + w / 2;
        if (-d.y - nodeHeight / 2 < minY) minY = -d.y - nodeHeight / 2;
        if (-d.y + nodeHeight / 2 > maxY) maxY = -d.y + nodeHeight / 2;
      });

      const treeWidth = maxX - minX;
      const treeHeight = maxY - minY;
      const viewWidth = window.innerWidth;
      const viewHeight = window.innerHeight;

      const scaleX = (viewWidth * 0.85) / treeWidth;
      const scaleY = (viewHeight * 0.75) / treeHeight;
      const targetScale = Math.min(Math.max(Math.min(scaleX, scaleY), 0.1), 1.0);

      const centerX = (minX + maxX) / 2;
      const centerY = (minY + maxY) / 2;

      const translateX = viewWidth / 2 - centerX * targetScale;
      const translateY = viewHeight / 2 - centerY * targetScale;

      svg.transition()
        .duration(750)
        .call(zoom.transform, d3.zoomIdentity.translate(translateX, translateY).scale(targetScale));
    }

    renderTree(maxGenFound);

    if (genSelect.node() && !genSelect.selectAll("option").empty()) {
      genSelect.on("change", function() {
        const selectedGen = parseInt(this.value, 10);
        renderTree(selectedGen);
      });
    }

  })
  .catch(err => console.error("Erreur de chargement du fichier Excel :", err));
