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

def _build_filter_clause(filters):
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
        clauses.append("m.excavation_unit_id = ?")
        params.append(filters['excavation_unit_id'])

    if filters.get('text_filter'):
        # Filtro genérico de texto (ajustado para Material)
        text = f"%{filters['text_filter']}%"
        clauses.append("(m.material_type LIKE ? OR m.material_description LIKE ? OR m.notes LIKE ?)")
        params.extend([text, text, text])

    where_clause = " AND ".join(clauses)
    if where_clause:
        where_clause = "WHERE " + where_clause
    
    return where_clause, params

def plot_material_types(db, filters, canvas, settings=None):
    """Gera gráfico de barras por tipo de material."""
    where, params = _build_filter_clause(filters)
    
    query = f"""
        SELECT m.material_type, COUNT(*) as count
        FROM Material m
        JOIN ExcavationUnit u ON m.excavation_unit_id = u.id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY m.material_type
        ORDER BY count DESC
        LIMIT 20
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 6))

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [row[0] if row[0] else "Indeterminado" for row in data]
    counts = [row[1] for row in data]

    bar_width = settings.get('bar_width', 0.8) if settings else 0.8
    canvas.axes.bar(labels, counts, color='#4a90e2', width=bar_width)
    canvas.axes.set_title("Distribuição por Tipo de Material")
    canvas.axes.set_ylabel("Contagem")
    canvas.axes.set_xlabel("Tipo de Material")
    canvas.axes.tick_params(axis='x', rotation=45)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_material_descriptions(db, filters, canvas, settings=None):
    """Gera gráfico de distribuição por descrição de material."""
    where, params = _build_filter_clause(filters)

    query = f"""
        SELECT m.material_description, COUNT(*) as count
        FROM Material m
        JOIN ExcavationUnit u ON m.excavation_unit_id = u.id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where}
        GROUP BY m.material_description
        ORDER BY count DESC
        LIMIT 20
    """

    data = db.execute_query(query, tuple(params))

    canvas.axes.clear()
    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 6))

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [row[0] if row[0] else "Indeterminado" for row in data]
    counts = [row[1] for row in data]

    bar_width = settings.get('bar_width', 0.8) if settings else 0.8
    canvas.axes.bar(labels, counts, color='#50c878', width=bar_width)
    canvas.axes.set_title("Distribuição por Descrição de Material")
    canvas.axes.set_ylabel("Contagem")
    canvas.axes.tick_params(axis='x', rotation=45)
    canvas.fig.tight_layout()
    canvas.draw()

def plot_material_quantities(db, filters, canvas, settings=None):
    """Gera gráfico de distribuição por quantidade de material."""
    where, params = _build_filter_clause(filters)

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
    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 6))

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    labels = [str(row[0]) if row[0] is not None else "N/A" for row in data]
    counts = [row[1] for row in data]

    bar_width = settings.get('bar_width', 0.8) if settings else 0.8
    canvas.axes.bar(labels, counts, color='#e74c3c', width=bar_width)
    canvas.axes.set_title("Distribuição por Quantidade de Material")
    canvas.axes.set_ylabel("Contagem")
    canvas.axes.set_xlabel("Quantidade")
    canvas.fig.tight_layout()
    canvas.draw()

def plot_unit_counts(db, filters, canvas, settings=None):
    """Gera gráfico de contagem de materiais por unidade de escavação."""
    # Filtros para unidade são ligeiramente diferentes na estrutura WHERE se filtrarmos a tabela Unit diretamente,
    # mas aqui queremos contar Specimens POR Unit.
    
    # Reutiliza a lógica, mas o GROUP BY é na Unidade
    where, params = _build_filter_clause(filters)
    
    query = f"""
        SELECT u.name, COUNT(m.id) as count
        FROM ExcavationUnit u
        LEFT JOIN Material m ON u.id = m.excavation_unit_id
        JOIN Assemblage a ON u.assemblage_id = a.id
        JOIN Collection c ON a.collection_id = c.id
        {where.replace('m.excavation_unit_id', 'u.id') if where else ''} 
        GROUP BY u.name
        ORDER BY count DESC
        LIMIT 20
    """
    # Nota: O replace acima é um ajuste rápido pois _build_filter_clause assume alias 's' para specimen linkado a unit.
    # Se o filtro vier de 'excavation_unit_id', ele usa 's.excavation_unit_id'. 
    # Para contar por unidade, precisamos garantir que o filtro de unidade (se houver) se aplique a 'u.id'.
    
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
        LIMIT 20
    """

    data = db.execute_query(query, tuple(q_params))

    canvas.axes.clear()
    if settings:
        canvas.fig.set_size_inches(settings.get('width', 10), settings.get('height', 6))

    if not data:
        canvas.axes.text(0.5, 0.5, "Sem dados para exibir", ha='center')
        canvas.draw()
        return

    units = [row[0] for row in data]
    counts = [row[1] for row in data]

    bar_width = settings.get('bar_width', 0.8) if settings else 0.8
    canvas.axes.bar(units, counts, color='#9b59b6', width=bar_width)
    canvas.axes.set_title("Densidade de Materiais por Unidade")
    canvas.axes.set_ylabel("Total de Materiais")
    canvas.axes.tick_params(axis='x', rotation=45)
    canvas.fig.tight_layout()
    canvas.draw()
