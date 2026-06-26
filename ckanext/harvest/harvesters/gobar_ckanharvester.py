from __future__ import absolute_import

import json
import logging
import os
import unicodedata
import uuid

from ckan import model
from ckanext.harvest.model import HarvestObject

from .ckanharvester import CKANHarvester

log = logging.getLogger(__name__)


def _normalize_str(s):
    """Minúsculas y sin signos diacríticos para matching de nombres geográficos."""
    s = unicodedata.normalize('NFD', s.lower())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')

#TODO: una vez que se implemente script para actualización automática de status, reveer el hardcode de 'status'
#TODO: una vez que se implementen valores de hvdCategory, cambiar valor default


class GobArCKANHarvester(CKANHarvester):
    """
    Harvester para instancias CKAN remotas con transformaciones del perfil datos.gob.ar.

    Hereda toda la infraestructura de CKANHarvester (gather/fetch/import, búsqueda
    paginada, manejo de orgs y grupos) y sobreescribe únicamente modify_package_dict
    y modify_resource_dict para aplicar las reglas del perfil de metadatos local.

    Transformaciones aplicadas a ckan remotos V1:
    - Renombrado de campos core que vienen con nombres distintos en catálogos v1
      (author → dataset_publisher_name, etc.)
    - Promoción de extras DCAT a campos de primer nivel del esquema
    - Traducción de superTheme a URIs de vocabularios controlados; accrualPeriodicity se preserva tal como viene
    - Normalización de spatial: si es GeoJSON válido se preserva; si son códigos de
      provincia se traducen a URIs; en cualquier otro caso se vacía
    - Normalización de temporal: convierte sufijo Z a +00:00 (para evitar problemas de isoformat en los
      validadores) y parte intervalos ISO 8601 (start/end) en temporal_start y temporal_end
    - Defaults para dataset_status, dataset_hvdCategory y dataset_language
    - Generación de IDs locales deterministas con UUID5 para garantizar que el mismo
      dataset remoto siempre produzca el mismo ID local entre corridas del harvester
    - Limpieza de recursos: whitelist de campos, resolución de mimetype a URI IANA
      y asignación de category según el formato del recurso
    """

    def info(self):
        """
        Retorna los metadatos de registro del harvester para CKAN.

        El campo 'name' debe coincidir con el entry point declarado en pyproject.toml
        (gobar_ckan_harvester). CKAN usa este dict para mostrar el harvester en la
        interfaz de administración y para identificarlo internamente.
        """
        return {
            'name': 'gobar_ckan',
            'title': 'GobAr CKAN',
            'description': (
                'Harvests remote CKAN instances applying datos.gob.ar '
                'metadata profile transformations.'
            ),
            'form_config_interface': 'Text',
        }

    # ------------------------------------------------------ gather_stage --

    def gather_stage(self, harvest_job):
        """Extiende el gather_stage del CKANHarvester para incluir datasets sin cambios.

        El CKANHarvester optimiza el gather pidiendo solo los datasets modificados
        desde el último job exitoso. Cuando no hubo cambios devuelve [] y el job
        termina sin objetos, dejando el contador "not modified" siempre en 0.

        Este override detecta qué datasets ya existían (current=True) y no fueron
        incluidos en el gather optimizado, y crea HarvestObjects para ellos copiando
        el contenido del job anterior. El import_stage los procesará y los detectará
        como 'unchanged' vía la comparación de metadata_modified.
        """
        object_ids = super().gather_stage(harvest_job)

        if object_ids is None:
            return object_ids

        guids_gathered = set()
        for obj_id in object_ids:
            obj = HarvestObject.get(obj_id)
            if obj:
                guids_gathered.add(obj.guid)

        q = model.Session.query(HarvestObject).filter(
            HarvestObject.harvest_source_id == harvest_job.source.id,
            HarvestObject.current == True,  # noqa: E712
        )
        if guids_gathered:
            q = q.filter(HarvestObject.guid.notin_(guids_gathered))
        unchanged_objects = q.all()

        for prev_obj in unchanged_objects:
            new_obj = HarvestObject(
                guid=prev_obj.guid,
                job=harvest_job,
                content=prev_obj.content,
            )
            new_obj.save()
            object_ids.append(new_obj.id)

        return object_ids

    # ------------------------------------------------------ import_stage --

    def import_stage(self, harvest_object):
        result = super().import_stage(harvest_object)
        if result == 'unchanged':
            harvest_object.report_status = 'not modified'
            harvest_object.save()
        return result

    # --------------------------------------------------------- fields_options --

    @property
    def fields_options(self):
        """
        Carga y cachea el archivo assets/fields_options.json.

        El JSON contiene los diccionarios de traducción de valores controlados
        usados en las transformaciones del perfil:
          - superthemes: mapeo de siglas DCAT a URIs EU Publications
          - province_codes: mapeo de códigos numéricos de provincia argentina a URIs
          - mimetypes: mapeo de media types IANA a sus URIs canónicas
          - category_formats: mapeo de extensiones de archivo a URIs de tipo DCMI

        Se lee una sola vez por instancia del harvester y se guarda en _field_options.
        Retorna {} si el archivo no existe o no puede parsearse, permitiendo que el
        harvester siga funcionando sin transformaciones de vocabulario controlado.
        """
        if not hasattr(self, '_field_options'):
            self._field_options = None
        if self._field_options:
            return self._field_options
        path = os.path.join(os.path.dirname(__file__), '../assets/fields_options.json')
        try:
            with open(path, 'r') as f:
                self._field_options = json.load(f)
        except Exception as e:
            log.warning('Error cargando fields_options.json en %s: %s', path, e)
            self._field_options = {}
        return self._field_options

    def get_field_options(self, field_name):
        """
        Retorna el dict de opciones para field_name desde fields_options.json.

        Es el punto de acceso uniforme a todos los vocabularios controlados.
        Retorna {} si field_name no existe en el JSON, lo que permite usar
        .get() sobre el resultado sin riesgo de AttributeError.
        """
        return self.fields_options.get(field_name, {})

    # ---------------------------------------------------- modify_package_dict --

    def modify_package_dict(self, package_dict, harvest_object):
        """
        Aplica las transformaciones del perfil datos.gob.ar al package_dict.

        Recibe el package_dict ya armado por el import_stage de CKANHarvester
        (con owner_org asignado, extras procesados y recursos sin url_type ni
        revision_id) y lo transforma in-place antes de que _create_or_update_package
        lo persista en CKAN.

        Transformaciones en orden:

        1. Renombrado de campos v1 (ANDINO_V1_DATASET_CKAN_MAP):
           Catálogos CKAN más viejos usan 'author' y 'author_email' como nombres
           de campo. Se renombran a los nombres del esquema actual para evitar que
           queden huérfanos en el package_dict. El campo 'notes' (descripción) se
           preserva tal cual ya que es el campo canónico del esquema.

        2. Promoción de extras a campos de primer nivel (EXTRAS_TO_FIELDS):
           El CKAN remoto puede enviar campos DCAT como extras en lugar de campos
           raíz. Se extrae el valor de cada extra conocido y se asigna al campo
           correspondiente, siempre que el campo destino no tenga ya un valor.
           Las claves en SCHEMA_COLLISION_KEYS (modified, spatial, temporal) se
           descartan del listado de extras porque colisionan con campos de esquema
           de CKAN y causarían errores de validación si se dejaran como extras.

        3. dataset_accrualPeriodicity:
           Se preserva el valor tal como viene, sin mapeo.

        4. dataset_superTheme:
           Acepta el valor como string JSON, lista Python o string simple. Traduce
           cada sigla temática (ej. 'AGRI') a su URI EU Publications. Si el campo
           está vacío o no se puede parsear, asigna ['Sin tema'] como default.

        5. spatial:
           Si el valor es un GeoJSON válido (dict con clave 'type'), se preserva
           sin cambios. Si no, intenta interpretar los valores como códigos numéricos
           de provincia argentina y los traduce a URIs del vocabulario de territorio.
           El resultado se guarda en spatial_uri y spatial se vacía. Si no hay
           coincidencias con ningún código conocido, ambos campos se vacían.

        6. temporal_start:
           Normaliza el sufijo 'Z' a '+00:00' para compatibilidad con parsers ISO
           8601 estrictos. Si el valor contiene '/' (intervalo ISO 8601), lo parte
           en temporal_start y temporal_end. Si no hay temporal_end, lo iguala a
           temporal_start para que el campo siempre tenga un valor de cierre.

        7. Defaults de campos opcionales:
           - dataset_status: 'Completed' (ADMS) si está vacío
           - dataset_hvdCategory: 'no aplica' si está vacío
           - dataset_issued / dataset_modified: fallback a metadata_created /
             metadata_modified del paquete CKAN si están vacíos
           - dataset_language: siempre se sobreescribe con la URI del español
             (EU Publications Office), independientemente del valor remoto
           - dataset_theme: siempre 'Tema específico 1' hasta que el tesauro
             local esté disponible

        8. ID del dataset con UUID5:
           Genera un UUID5 determinista usando NAMESPACE_DNS y la semilla
           "owner_org/raw_id". Dado que owner_org y raw_id son estables entre
           corridas del harvester, el mismo dataset remoto siempre produce el
           mismo UUID local, permitiendo que _create_or_update_package identifique
           si debe crear o actualizar sin necesidad de estado externo.
           El ID original del CKAN remoto se preserva en 'original_identifier'
           para trazabilidad. Si no hay raw_id se genera un UUID4 aleatorio.

        9. ID de recursos con UUID5:
           Para cada recurso aplica la misma lógica pero con NAMESPACE_URL y la
           semilla "pkg_id/raw_res_id", donde pkg_id es el UUID5 del dataset ya
           calculado. El ID original del recurso remoto se preserva en
           'original_identifier'. Luego cada recurso pasa por modify_resource_dict
           para su limpieza y normalización final.

        Retorna el package_dict modificado.
        """

        # 1. Renombrar campos core que vienen con nombres distintos en CKAN v1
        ANDINO_V1_DATASET_CKAN_MAP = {
            'author':       'dataset_publisher_name',
            'author_email': 'dataset_publisher_mbox',
        }
        for src, dst in ANDINO_V1_DATASET_CKAN_MAP.items():
            if src in package_dict:
                package_dict[dst] = package_dict.pop(src)

        # 2. Promover extras relevantes a campos de primer nivel
        EXTRAS_TO_FIELDS = {
            'accrualPeriodicity': 'dataset_accrualPeriodicity',
            'superTheme':         'dataset_superTheme',
            'modified':           'dataset_modified',
            'Modificado':         'dataset_modified',
            'Publicado':          'dataset_issued',
            'spatial':            'spatial',
            'language':           'dataset_language',
            'temporal':           'temporal_start',
        }
        SCHEMA_COLLISION_KEYS = {'modified', 'Modificado', 'spatial', 'temporal'}

        remaining_extras = []
        for extra in package_dict.get('extras', []):
            key = extra.get('key', '')
            value = extra.get('value', '')
            if key in EXTRAS_TO_FIELDS:
                field_name = EXTRAS_TO_FIELDS[key]
                if not package_dict.get(field_name):
                    package_dict[field_name] = value
            elif key not in SCHEMA_COLLISION_KEYS:
                remaining_extras.append(extra)
        package_dict['extras'] = remaining_extras

        # 3. dataset_accrualPeriodicity → default si vacío
        if not package_dict.get('dataset_accrualPeriodicity'):
            package_dict['dataset_accrualPeriodicity'] = 'sin especificar'

        # 4. dataset_superTheme → lista de URIs controladas
        st_value = package_dict.get('dataset_superTheme')
        if st_value:
            try:
                st_list = json.loads(st_value) if isinstance(st_value, str) else st_value
            except (ValueError, TypeError):
                st_list = [st_value]
            superthemes = self.get_field_options('superthemes')
            package_dict['dataset_superTheme'] = [
                superthemes.get(item, item) for item in st_list
            ]
        else:
            package_dict['dataset_superTheme'] = ['Sin tema']

        # 5. spatial → GeoJSON (se preserva) o códigos de provincia → URIs
        spatial = package_dict.get('spatial', '')
        is_geojson = False
        if spatial:
            try:
                parsed = json.loads(spatial)
                if isinstance(parsed, dict) and parsed.get('type'):
                    is_geojson = True
            except (ValueError, TypeError):
                pass

        if not is_geojson:
            if isinstance(spatial, list):
                candidates = [str(v).strip() for v in spatial if str(v).strip()]
            elif spatial:
                try:
                    parsed = json.loads(spatial)
                    candidates = parsed if isinstance(parsed, list) else [str(parsed).strip()]
                except (ValueError, TypeError):
                    candidates = [v.strip() for v in str(spatial).split(',') if v.strip()]
            else:
                candidates = []

            province_codes = self.get_field_options('province_codes')
            matched_uris = [province_codes[c] for c in candidates if c in province_codes]

            if matched_uris:
                package_dict['spatial_uri'] = matched_uris
                package_dict['spatial'] = ''
            else:
                package_dict['spatial_uri'] = ''
                package_dict['spatial'] = ''

        # 5b. Fallback: buscar nombre de provincia/región en el título si spatial_uri sigue vacío
        if not package_dict.get('spatial_uri'):
            title_norm = _normalize_str(package_dict.get('title', ''))
            province_names = self.get_field_options('province_names')
            matches = [(name, uri) for name, uri in province_names.items() if name in title_norm]
            matched_name_set = {name for name, _ in matches}
            matched_uris = list({
                uri for name, uri in matches
                if not any(name != other and name in other for other in matched_name_set)
            })
            if matched_uris:
                package_dict['spatial_uri'] = matched_uris

        # 6. temporal_start → normalizar sufijo Z y partir intervalo ISO 8601
        temporal_start = str(package_dict.get('temporal_start', '') or '')
        temporal_start = temporal_start.replace('Z', '+00:00')

        if temporal_start:
            if '/' in temporal_start:
                parts = temporal_start.split('/', 1)
                package_dict['temporal_start'] = parts[0]
                package_dict['temporal_end'] = parts[1]
            else:
                package_dict['temporal_start'] = temporal_start
                if not package_dict.get('temporal_end'):
                    package_dict['temporal_end'] = temporal_start

        # 7. Defaults de campos opcionales
        if not package_dict.get('dataset_status'):
            title = package_dict.get('title', '')
            if 'discontinuad' in title.lower():
                package_dict['dataset_status'] = 'http://purl.org/adms/status/Withdrawn'
            else:
                package_dict['dataset_status'] = 'http://purl.org/adms/status/Completed'

        if not package_dict.get('dataset_hvdCategory'):
            package_dict['dataset_hvdCategory'] = 'no aplica'

        if not package_dict.get('dataset_issued'):
            package_dict['dataset_issued'] = package_dict.get('metadata_created', '')
        if not package_dict.get('dataset_modified'):
            package_dict['dataset_modified'] = package_dict.get('metadata_modified', '')

        package_dict['dataset_language'] = (
            'http://publications.europa.eu/resource/authority/language/SPA'
        )

        # TODO: sobreescribe dataset_theme hasta que el tesauro esté en revisión
        package_dict['dataset_theme'] = 'Tema específico 1'

        # 8. UUID5 para el dataset: UUID5(NAMESPACE_DNS, "owner_org/raw_id")
        raw_id = str(package_dict.get('id', '') or '').strip()
        if raw_id:
            seed = '%s/%s' % (package_dict.get('owner_org', ''), raw_id)
            package_dict['id'] = str(uuid.uuid5(uuid.NAMESPACE_DNS, seed))
        else:
            package_dict['id'] = str(uuid.uuid4())
        if raw_id and raw_id != package_dict['id']:
            package_dict['original_identifier'] = raw_id

        pkg_id = package_dict['id']

        # 9. UUID5 para cada recurso: UUID5(NAMESPACE_URL, "pkg_id/raw_res_id")
        for resource in package_dict.get('resources', []):
            raw_res_id = str(resource.get('id', '') or '').strip()
            if raw_res_id:
                seed = '%s/%s' % (pkg_id, raw_res_id)
                resource['id'] = str(uuid.uuid5(uuid.NAMESPACE_URL, seed))
            else:
                resource['id'] = str(uuid.uuid4())
            if raw_res_id and raw_res_id != resource['id']:
                resource['original_identifier'] = raw_res_id

        package_dict['resources'] = [
            self.modify_resource_dict(resource)
            for resource in package_dict.get('resources', [])
        ]

        return package_dict

    # ---------------------------------------------------- modify_resource_dict --

    def modify_resource_dict(self, resource):
        """
        Limpia y normaliza un dict de recurso antes de persistirlo en CKAN.

        Construye un nuevo dict (clean_resource) con una whitelist explícita de
        campos, descartando cualquier campo extra que venga del CKAN remoto que
        no forme parte del esquema de distribuciones de datos.gob.ar. Los campos
        geoespaciales opcionales (character_set, scale, projection, iso19115_url,
        wfs_url) se inicializan en '' para garantizar que siempre estén presentes
        en el esquema aunque el recurso remoto no los traiga.

        Campos preservados del recurso original:
        - id: UUID5 ya asignado por modify_package_dict
        - name, url, description, format: campos core de CKAN
        - mimetype: se toma de 'mimetype' o 'mediaType' (compatibilidad DCAT)
        - last_modified: se toma de 'modified' o 'last_modified'
        - created: se toma de 'created' o 'issued'
        - original_identifier: UUID remoto guardado para trazabilidad

        Transformaciones aplicadas:

        mimetype:
            Traduce el media type crudo (ej. 'text/csv') a su URI canónica IANA
            (ej. 'https://www.iana.org/assignments/media-types/text/csv') usando
            el diccionario 'mimetypes' de fields_options.json. Si el valor no
            está en el diccionario, asigna 'other'.

        category:
            Determina el tipo DCMI del recurso (Dataset, Text, Image, etc.)
            buscando el formato del recurso (lowercased) en las listas de
            extensiones del diccionario 'category_formats' de fields_options.json.
            Retorna la URI DCMI correspondiente o 'other' si no hay coincidencia.

        Retorna el clean_resource normalizado.
        """
        clean_resource = {
            'id':                  resource.get('id', ''),
            'name':                resource.get('name', ''),
            'url':                 resource.get('url', ''),
            'description':         resource.get('description', ''),
            'format':              resource.get('format', ''),
            'mimetype':            resource.get('mimetype', '') or resource.get('mediaType', ''),
            'character_set':       '',
            'scale':               '',
            'projection':          '',
            'iso19115_url':        '',
            'wfs_url':             '',
            'last_modified':       resource.get('modified', '') or resource.get('last_modified', ''),
            'created':             resource.get('created', '') or resource.get('issued', ''),
            'original_identifier': resource.get('original_identifier', ''),
        }

        # Resolver mimetype → URI IANA solo si no es ya una URI
        media_type = clean_resource['mimetype']
        if not media_type:
            fmt = clean_resource['format'].lower()
            media_type = self.get_field_options('format_to_mediatype').get(fmt, '')
        if media_type.startswith('http'):
            clean_resource['mimetype'] = media_type
        else:
            clean_resource['mimetype'] = (
                self.get_field_options('mimetypes').get(media_type) or 'other'
            )

        # Resolver category según formato solo si no viene ya asignada
        existing_category = resource.get('category', '')
        if existing_category:
            clean_resource['category'] = existing_category
        else:
            resource_format = clean_resource['format'].lower()
            clean_resource['category'] = next(
                (k for k, v in self.get_field_options('category_formats').items()
                 if resource_format in v),
                'other'
            )

        return clean_resource
