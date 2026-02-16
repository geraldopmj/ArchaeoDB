import matplotlib
# Garante o uso do backend correto para PySide6
matplotlib.use('QtAgg')

from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

class MplCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)
        self.setParent(parent)

def _build_filter_clause(filters, unit_col='m.excavation_unit_id', text_cols=None):
    """Constrói a cláusula WHERE e os parâmetros baseados nos filtros hierárquicos."""
    clauses = []
    params = []
    
    # Mapeamento de filtros para colunas e tabelas
    # Estrutura: Material -> ExcavationUnit -> Assemblage -> Collection -> Site
    
    if filters.get('site_id'):
        clauses.append("c.site_id = ?")
        params.append(filters['site_id'])
        
    if filters.get('collection_id'):
        clauses.append("a.collection_id = ?")
        params.append(filters['collection_id'])
        
    if filters.get('assemblage_id'):
        clauses.append("u.assemblage_id = ?")
        params.append(filters['assemblage_id'])
        
    if filters.get('excavation_unit_id'):
        clauses.append(f"{unit_col} = ?")
        params.append(filters['excavation_unit_id'])

    if filters.get('text_filter') and text_cols:
        text = f"%{filters['text_filter']}%"
        or_clauses = " OR ".join([f"{col} LIKE ?" for col in text_cols])
        clauses.append(f"({or_clauses})")
        params.extend([text] * len(text_cols))

    where_clause = " AND ".join(clauses)
    if where_clause:
        where_clause = "WHERE " + where_clause
    
    return where_clause, params

def plot_material_types(db, filters, canvas, settings=None):
    """Gera gráfico de barras por tipo de material."""
    where, params = _build_filter_clause(filters, unit_col='m.excavation_unit_id')
    
    query = f"""
        SELECT m.material_type, COUNT(*) as count
        FROM Material m
        JOIN ExcavationUnit u ON m.excavation_unit_id = u.id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY m.material_type
        ORDER BY count DESC
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    
    title = "Distribuição por Tipo de Material"
    xlabel = "Tipo de Material"
    ylabel = "Contagem"
    bar_width = 0.8

    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 5.6))
        canvas.setMinimumSize(int(settings.get('width', 10) * canvas.fig.dpi), int(settings.get('height', 5.6) * canvas.fig.dpi))
        if settings.get('title'): title = settings.get('title')
        if settings.get('xlabel'): xlabel = settings.get('xlabel')
        if settings.get('ylabel'): ylabel = settings.get('ylabel')
        bar_width = settings.get('bar_width', 0.8)

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [row[0] if row[0] else "Indeterminado" for row in data]
    counts = [row[1] for row in data]

    canvas.axes.bar(labels, counts, color='#4a90e2', width=bar_width)
    
    if settings and settings.get('show_labels'):
        for i, v in enumerate(counts):
            canvas.axes.text(i, v, str(v), ha='center', va='bottom', fontsize=8)

    canvas.axes.set_title(title)
    canvas.axes.set_ylabel(ylabel)
    canvas.axes.set_xlabel(xlabel)
    
    rotation = 90 if settings and settings.get('vertical_xlabel') else 45
    canvas.axes.tick_params(axis='x', rotation=rotation)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_material_descriptions(db, filters, canvas, settings=None):
    """Gera gráfico de distribuição por descrição de material."""
    where, params = _build_filter_clause(filters, unit_col='m.excavation_unit_id')

    query = f"""
        SELECT m.material_description, COUNT(*) as count
        FROM Material m
        JOIN ExcavationUnit u ON m.excavation_unit_id = u.id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY m.material_description
        ORDER BY count DESC
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    
    title = "Distribuição por Descrição de Material"
    xlabel = "Descrição"
    ylabel = "Contagem"
    bar_width = 0.8

    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 5.6))
        canvas.setMinimumSize(int(settings.get('width', 10) * canvas.fig.dpi), int(settings.get('height', 5.6) * canvas.fig.dpi))
        if settings.get('title'): title = settings.get('title')
        if settings.get('xlabel'): xlabel = settings.get('xlabel')
        if settings.get('ylabel'): ylabel = settings.get('ylabel')
        bar_width = settings.get('bar_width', 0.8)

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [row[0] if row[0] else "Indeterminado" for row in data]
    counts = [row[1] for row in data]

    canvas.axes.bar(labels, counts, color='#50c878', width=bar_width)

    if settings and settings.get('show_labels'):
        for i, v in enumerate(counts):
            canvas.axes.text(i, v, str(v), ha='center', va='bottom', fontsize=8)

    canvas.axes.set_title(title)
    canvas.axes.set_ylabel(ylabel)
    canvas.axes.set_xlabel(xlabel)
    
    rotation = 90 if settings and settings.get('vertical_xlabel') else 45
    canvas.axes.tick_params(axis='x', rotation=rotation)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_material_quantities(db, filters, canvas, settings=None):
    """Gera gráfico de distribuição por quantidade de material."""
    where, params = _build_filter_clause(filters, unit_col='m.excavation_unit_id')

    query = f"""
        SELECT m.quantity, COUNT(*) as count
        FROM Material m
        JOIN ExcavationUnit u ON m.excavation_unit_id = u.id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY m.quantity
        ORDER BY count DESC
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    
    title = "Distribuição por Quantidade de Material"
    xlabel = "Quantidade"
    ylabel = "Contagem"
    bar_width = 0.8

    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 5.6))
        canvas.setMinimumSize(int(settings.get('width', 10) * canvas.fig.dpi), int(settings.get('height', 5.6) * canvas.fig.dpi))
        if settings.get('title'): title = settings.get('title')
        if settings.get('xlabel'): xlabel = settings.get('xlabel')
        if settings.get('ylabel'): ylabel = settings.get('ylabel')
        bar_width = settings.get('bar_width', 0.8)

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [str(row[0]) if row[0] is not None else "N/A" for row in data]
    counts = [row[1] for row in data]

    canvas.axes.bar(labels, counts, color='#e74c3c', width=bar_width)

    if settings and settings.get('show_labels'):
        for i, v in enumerate(counts):
            canvas.axes.text(i, v, str(v), ha='center', va='bottom', fontsize=8)

    canvas.axes.set_title(title)
    canvas.axes.set_ylabel(ylabel)
    canvas.axes.set_xlabel(xlabel)
    
    rotation = 90 if settings and settings.get('vertical_xlabel') else 45
    canvas.axes.tick_params(axis='x', rotation=rotation)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_unit_counts(db, filters, canvas, settings=None):
    """Gera gráfico de contagem de materiais por unidade de escavação."""
    # Ajuste manual simples para garantir integridade dos filtros de unidade na query de agregação
    clauses = []
    q_params = []
    if filters.get('site_id'):
        clauses.append("c.site_id = ?")
        q_params.append(filters['site_id'])
    if filters.get('collection_id'):
        clauses.append("a.collection_id = ?")
        q_params.append(filters['collection_id'])
    if filters.get('assemblage_id'):
        clauses.append("u.assemblage_id = ?")
        q_params.append(filters['assemblage_id'])
    if filters.get('excavation_unit_id'):
        clauses.append("u.id = ?")
        q_params.append(filters['excavation_unit_id'])
        
    where_clause = " AND ".join(clauses)
    if where_clause:
        where_clause = "WHERE " + where_clause

    query = f"""
        SELECT u.name, COUNT(m.id) as count
        FROM ExcavationUnit u
        LEFT JOIN Material m ON u.id = m.excavation_unit_id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where_clause}
        GROUP BY u.id
        ORDER BY count DESC
    """

    data = db.execute_query(query, tuple(q_params))

    canvas.axes.clear()
    
    title = "Densidade de Materiais por Unidade"
    xlabel = "Unidade"
    ylabel = "Total de Materiais"
    bar_width = 0.8

    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 5.6))
        canvas.setMinimumSize(int(settings.get('width', 10) * canvas.fig.dpi), int(settings.get('height', 5.6) * canvas.fig.dpi))
        if settings.get('title'): title = settings.get('title')
        if settings.get('xlabel'): xlabel = settings.get('xlabel')
        if settings.get('ylabel'): ylabel = settings.get('ylabel')
        bar_width = settings.get('bar_width', 0.8)

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    units = [row[0] for row in data]
    counts = [row[1] for row in data]

    canvas.axes.bar(units, counts, color='#9b59b6', width=bar_width)

    if settings and settings.get('show_labels'):
        for i, v in enumerate(counts):
            canvas.axes.text(i, v, str(v), ha='center', va='bottom', fontsize=8)

    canvas.axes.set_title(title)
    canvas.axes.set_ylabel(ylabel)
    canvas.axes.set_xlabel(xlabel)
    
    rotation = 90 if settings and settings.get('vertical_xlabel') else 45
    canvas.axes.tick_params(axis='x', rotation=rotation)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_specimen_nisp(db, filters, canvas, settings=None):
    """Gera gráfico NISP (Number of Identified Specimens) por Taxon."""
    # Usa 's' como alias para Specimen e 's.excavation_unit_id' para filtro de unidade
    where, params = _build_filter_clause(filters, unit_col='s.excavation_unit_id')
    
    query = f"""
        SELECT s.taxon, COUNT(*) as count
        FROM Specimen s
        LEFT JOIN ExcavationUnit u ON s.excavation_unit_id = u.id
        LEFT JOIN Assemblage a ON u.assemblage_id = a.id
        LEFT JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY s.taxon
        ORDER BY count DESC
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    
    title = "NISP por Taxon"
    xlabel = "Taxon"
    ylabel = "NISP"
    bar_width = 0.8

    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 5.6))
        canvas.setMinimumSize(int(settings.get('width', 10) * canvas.fig.dpi), int(settings.get('height', 5.6) * canvas.fig.dpi))
        if settings.get('title'): title = settings.get('title')
        if settings.get('xlabel'): xlabel = settings.get('xlabel')
        if settings.get('ylabel'): ylabel = settings.get('ylabel')
        bar_width = settings.get('bar_width', 0.8)

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [row[0] if row[0] else "Indeterminado" for row in data]
    counts = [row[1] for row in data]

    canvas.axes.bar(labels, counts, color='#e67e22', width=bar_width)

    if settings and settings.get('show_labels'):
        for i, v in enumerate(counts):
            canvas.axes.text(i, v, str(v), ha='center', va='bottom', fontsize=8)

    canvas.axes.set_title(title)
    canvas.axes.set_ylabel(ylabel)
    canvas.axes.set_xlabel(xlabel)
    
    rotation = 90 if settings and settings.get('vertical_xlabel') else 45
    canvas.axes.tick_params(axis='x', rotation=rotation)
    canvas.fig.tight_layout()
    canvas.draw()
