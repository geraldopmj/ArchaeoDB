# models.py (Definição das Tabelas)
from dataclasses import dataclass
from typing import Optional

@dataclass
class Site:
    id: Optional[int] = None
    name: Optional[str] = None
    state: Optional[str] = None
    city: Optional[str] = None
    location: Optional[str] = None
    number: Optional[int] = None
    longitude: Optional[float] = None
    latitude: Optional[float] = None

@dataclass
class Collection:
    id: Optional[int] = None
    site_id: Optional[int] = None
    name: Optional[str] = None
    longitude: Optional[float] = None
    latitude: Optional[float] = None
    
@dataclass
class Assemblage:
    id: Optional[int] = None
    collection_id: Optional[int] = None
    name: Optional[str] = None
    longitude: Optional[float] = None
    latitude: Optional[float] = None
    screenSize: Optional[str] = None
    
@dataclass
class Material:
    id: Optional[int] = None
    excavation_unit_id: Optional[int] = None
    level_id: Optional[int] = None
    uuid: Optional[str] = None
    field_serial: Optional[str] = None
    lab_serial: Optional[str] = None
    material_type: Optional[str] = None
    material_description: Optional[str] = None
    measurements: Optional[float] = None
    weight: Optional[float] = None
    quantity: Optional[int] = None
    x_coord: Optional[float] = None
    y_coord: Optional[float] = None
    z_coord: Optional[float] = None
    notes: Optional[str] = None
    photos: Optional[str] = None
    user: Optional[str] = None
    cataloging_date: Optional[str] = None


@dataclass
class ExcavationUnit:
    id: Optional[int] = None
    assemblage_id: Optional[int] = None
    name: Optional[str] = None
    unit_type: Optional[str] = None
    size: Optional[str] = None
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    nivel_holotipo: Optional[str] = None
    profundidade_inicial: Optional[float] = None
    profundidade_final: Optional[float] = None
    camada_geologica: Optional[str] = None
    metodo_escavacao: Optional[str] = None
    metodo_peneiramento: Optional[str] = None
    responsavel_escavacao: Optional[str] = None
    data_inicio: Optional[str] = None
    data_conclusao: Optional[str] = None
    fotos_registro: Optional[str] = None
    observacao: Optional[str] = None

@dataclass
class Level:
    id: Optional[int] = None
    excavation_unit_id: Optional[int] = None
    level: Optional[str] = None
    start_depth: Optional[float] = None
    end_depth: Optional[float] = None
    color: Optional[str] = None
    texture: Optional[str] = None
    description: Optional[str] = None