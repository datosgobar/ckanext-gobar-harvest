from __future__ import absolute_import

import io
import json
import logging
import hashlib
import os
import uuid
import requests
import openpyxl

from ckan import model
from ckan.logic import ValidationError, get_action
from ckan.plugins import toolkit

from ckanext.harvest.model import HarvestObject, HarvestGatherError
from .base import HarvesterBase

log = logging.getLogger(__name__)

#TODO:Revisar todos los mapeos de campos y que no se pierda nada

# ---------------------------------------------------------------------------
# Mapeo de encabezados Excel → nombre de campo interno
# Los encabezados siguen la display_property del perfil DCAT-AP datos.gob.ar
# ---------------------------------------------------------------------------

DATASET_COLUMN_MAP_OLD = {
    "dataset_identifier":     "id",   # UUID clave primaria
    "dataset_title":          "dataset_title",
    "name":                   "name",
    "dataset_publisher_name":      "owner_org",
    "dataset_publisher_mbox": "dataset_publisher_mbox",
    "dataset_description":    "dataset_description",
    "dataset_issued":             "dataset_issued",
    "dataset_modified":           "dataset_modified",
    "adms:status":            "dataset_status",
    "dct:theme":              "dataset_superTheme",
    "dcat:theme":             "dataset_theme",
    "dataset_keywords":       "dataset_keywords",
    "dcatap:hdvCategory":     "dataset_hvdcategory",
    "dct:accrualPeriodicity": "dataset_accrualPeriodicity",
    "temporal_start":         "temporal_start",
    "temporal_end":           "temporal_end",
    "dct:spatial":            "spatial_uri",
    "spatial_coverage":       "spatial_coverage",
    "license_id":             "license_id",
    "dataset_source":         "dataset_source",
}

DATASET_COLUMN_MAP = {
    "dataset_identifier":     "id",   # UUID clave primaria
    "dataset_title":          "dataset_title",
    "name":                   "name",
    "dataset_publisher_name": "dataset_publisher_name",
    "dataset_publisher_mbox": "dataset_publisher_mbox",
    "dataset_description":    "dataset_description",
    "dataset_issued":             "dataset_issued",
    "dataset_modified":           "dataset_modified",
    "adms:status":            "dataset_status",
    "dataset_superTheme":              "dataset_superTheme",
    "dataset_theme":             "dataset_theme",
    "dataset_keywords":       "dataset_keywords",
    "dcatap:hdvCategory":     "dataset_hvdcategory",
    "dataset:accrualPeriodicity": "dataset_accrualPeriodicity",
    "temporal_start":         "temporal_start",
    "temporal_end":           "temporal_end",
    "dataset_spatial":            "dataset_spatial",
    "spatial_coverage":       "spatial_coverage",
    "license_id":             "license_id",
    "dataset_source":         "dataset_source",
}

RESOURCE_COLUMN_MAP = {
    "dataset_identifier":            "id",      # FK → Datasets
    "distribution_identifier":       "distribution_identifier",
    "url":                           "url",
    "name":                          "name",
    "description":                   "description",
    "format":                        "format",
    "dcat:mediaType":                "distribution_mediaType",
    "dct:type":                      "distribution_type",
    "cnt:characterEncoding":         "distribution_character_set",
    "distribution_scale":            "distribution_scale",
    "dct:conformsTo":                "distribution_projection",
    "distribution_iso19115_url":     "distribution_iso19115_url",
    "distribution_wfs_url":          "distribution_wfs_url",
}

DEFAULT_DATASET_SHEET      = "dataset"
DEFAULT_DISTRIBUTION_SHEET = "distribution"

#DATASET_REQUIRED      = {"dataset_title", "owner_org", "dataset_description"}
DATASET_REQUIRED      = { }
DISTRIBUTION_REQUIRED = {"url"}


class XLSXHarvesterClaude(HarvesterBase):
    """
    Harvester para archivos Excel (.xlsx) con el mismo comportamiento que
    CKANHarvester: mismas transformaciones de campos, mapeo de valores y
    defaults via modify_package_dict().

    Estructura esperada del archivo:
      - Pestaña "Datasets"      : una fila por dataset; encabezados = display_property DCAT-AP
      - Pestaña "Distributions" : una fila por recurso; incluye dataset_identifier (FK)

    Configuración opcional (JSON en el campo "Configuration" del formulario):
    {
        "dataset_sheet":      "Datasets",
        "distribution_sheet": "Distributions",
        "skip_rows":          0,
        "default_owner_org":  "mi-org",
        "default_tags":       [{"name": "tag1"}],
        "default_extras":     {"clave": "valor"},
        "override_extras":    false,
        "user_agent":         "mi-portal/1.0",
        "force_all":          false
    }
    """

    config = None

    # ------------------------------------------------------------------ info --

    def info(self):
        return {
            "name": "xlsx_claude",
            "title": "XLSX (Claude)",
            "description": (
                "Harvests remote XLSX files following the datos.gob.ar DCAT-AP schema. "
                "Applies the same field transformations as the CKAN harvester. "
                "Expects a 'Datasets' sheet and a 'Distributions' sheet."
            ),
            "form_config_interface": "Text",
        }

    # --------------------------------------------------------- validate_config --

    def validate_config(self, config_str):
        if not config_str:
            return config_str
        try:
            cfg = json.loads(config_str)
        except ValueError as e:
            raise ValueError("La configuración debe ser un JSON válido: %s" % e)

        for key in ("dataset_sheet", "distribution_sheet", "default_owner_org", "user_agent"):
            if key in cfg and not isinstance(cfg[key], str):
                raise ValueError("%s debe ser un string" % key)
        if "skip_rows" in cfg and not isinstance(cfg["skip_rows"], int):
            raise ValueError("skip_rows debe ser un entero")
        if "default_tags" in cfg:
            if not isinstance(cfg["default_tags"], list):
                raise ValueError("default_tags debe ser una lista")
            if cfg["default_tags"] and not isinstance(cfg["default_tags"][0], dict):
                raise ValueError("default_tags debe ser una lista de diccionarios")
        if "default_extras" in cfg and not isinstance(cfg["default_extras"], dict):
            raise ValueError("default_extras debe ser un diccionario")
        for key in ("read_only", "force_all", "override_extras"):
            if key in cfg and not isinstance(cfg[key], bool):
                raise ValueError("%s debe ser booleano" % key)

        return config_str

    # ------------------------------------------------------ get_original_url --

    def get_original_url(self, harvest_object_id):
        obj = HarvestObject.get(harvest_object_id)
        return obj.source.url if obj else None

    # ------------------------------------------------------- fields_options --
    # Exact copy from CKANHarvester to ensure identical behavior.

    @property
    def fields_options(self):
        if not hasattr(self, "_field_options"):
            self._field_options = None
        if self._field_options:
            return self._field_options
        path = os.path.join(os.path.dirname(__file__), "../assets/fields_options.json")
        try:
            with open(path, "r") as f:
                self._field_options = json.load(f)
        except Exception as e:
            log.warning("Error cargando fields_options.json en %s: %s", path, e)
            self._field_options = {}
        return self._field_options

    def get_field_options(self, field_name):
        return self.fields_options.get(field_name, {})

    # ---------------------------------------------------------- gather_stage --

    def gather_stage(self, harvest_job):
        log.info("XLSXHarvesterClaude gather_stage: %s", harvest_job.source.url)

        self._set_config(harvest_job.source.config)
        source_url = harvest_job.source.url.strip()

        try:
            wb = self._fetch_workbook(source_url)
        except Exception as e:
            self._save_gather_error("No se pudo descargar/leer el XLSX: %s" % e, harvest_job)
            return []

        skip       = self.config.get("skip_rows", 0)
        ds_sheet   = self.config.get("dataset_sheet", DEFAULT_DATASET_SHEET)
        dist_sheet = self.config.get("distribution_sheet", DEFAULT_DISTRIBUTION_SHEET)

        try:
            ds_rows   = self._read_sheet(wb, ds_sheet,   DATASET_COLUMN_MAP,  skip)
            dist_rows = self._read_sheet(wb, dist_sheet, RESOURCE_COLUMN_MAP, skip)
        except Exception as e:
            self._save_gather_error("Error leyendo hojas del XLSX: %s" % e, harvest_job)
            return []

        if not ds_rows:
            self._save_gather_error("La hoja de datasets está vacía", harvest_job)
            return []

        # Agrupar distribuciones por dataset_identifier (FK)
        dist_by_dataset = {}
        for d in dist_rows:
            ds_id = d.get("id", "").strip()
            if ds_id:
                dist_by_dataset.setdefault(ds_id, []).append(d)

        object_ids = []
        seen_guids = set()

        for i, row in enumerate(ds_rows):
            ds_id = row.get("id", "").strip()

            missing = [f for f in DATASET_REQUIRED if not row.get(f, "").strip()]
            if missing:
                log.warning("Fila %d omitida: faltan campos %s", i + 2, missing)
                continue

            guid = self._make_guid(ds_id, source_url, i, row)

            if guid in seen_guids:
                log.info("Dataset duplicado descartado: %s", guid)
                continue
            seen_guids.add(guid)

            row["_resources"] = dist_by_dataset.get(ds_id, [])
            content = json.dumps(row, default=str)

            obj = HarvestObject(guid=guid, job=harvest_job, content=content, extras=[])
            obj.save()
            object_ids.append(obj.id)

        log.info("XLSXHarvesterClaude: %d datasets recolectados", len(object_ids))
        return object_ids

    # ----------------------------------------------------------- fetch_stage --

    def fetch_stage(self, harvest_object):
        # El contenido ya fue guardado en gather_stage
        if not harvest_object.content:
            self._save_object_error(
                "El harvest object no tiene contenido", harvest_object, "Fetch"
            )
            return False
        return True

    # ---------------------------------------------------------- import_stage --

    def import_stage(self, harvest_object):
        log.debug("XLSXHarvesterClaude import_stage: %s", harvest_object.id)

        base_context = {
            "model": model,
            "session": model.Session,
            "user": self._get_user_name(),
        }

        if not harvest_object.content:
            self._save_object_error("Contenido vacío", harvest_object, "Import")
            return False

        self._set_config(harvest_object.job.source.config)

        try:
            row = json.loads(harvest_object.content)
        except ValueError as e:
            self._save_object_error("JSON inválido: %s" % e, harvest_object, "Import")
            return False

        try:
            package_dict = self._build_package_dict(row, harvest_object)
        except Exception as e:
            self._save_object_error(
                "Error construyendo package_dict: %s" % e, harvest_object, "Import"
            )
            return False

        # Misma lógica de default_tags que CKANHarvester
        default_tags = self.config.get("default_tags", [])
        if default_tags:
            package_dict.setdefault("tags", [])
            package_dict["tags"].extend(
                [t for t in default_tags if t not in package_dict["tags"]]
            )

        # Asignar organización de la fuente si no viene en el XLSX
        source_dataset = get_action("package_show")(
            base_context.copy(), {"id": harvest_object.source.id}
        )
        local_org = source_dataset.get("owner_org")
        if not package_dict.get("owner_org"):
            package_dict["owner_org"] = local_org

        # Misma lógica de default_extras que CKANHarvester
        default_extras = self.config.get("default_extras", {})
        if default_extras:
            override_extras = self.config.get("override_extras", False)
            existing_extras = {e["key"]: e for e in package_dict.get("extras", [])}
            for key, value in default_extras.items():
                if key in existing_extras and not override_extras:
                    continue
                if key in existing_extras:
                    package_dict["extras"].remove(existing_extras[key])
                if isinstance(value, str):
                    value = value.format(
                        harvest_source_id=harvest_object.job.source.id,
                        harvest_source_url=harvest_object.job.source.url.strip("/"),
                        harvest_source_title=harvest_object.job.source.title,
                        harvest_job_id=harvest_object.job.id,
                        harvest_object_id=harvest_object.id,
                        dataset_id=package_dict.get("id", ""),
                    )
                package_dict["extras"].append({"key": key, "value": value})

        # Aplicar las mismas transformaciones de campos que CKANHarvester
        log.error(f"así está el package_dict antes de la modificacion: {package_dict}")
        package_dict = self.modify_package_dict(package_dict, harvest_object)
        log.error(f"así esta devolviendo el package_dict 15MAY: {package_dict}")




        try:
            return self._create_or_update_package(
                package_dict, harvest_object, package_dict_form="package_show"
            )
        except ValidationError as e:
            self._save_object_error(
                "Package inválido GUID %s: %r" % (harvest_object.guid, e.error_dict),
                harvest_object, "Import",
            )
        except Exception as e:
            self._save_object_error("%s" % e, harvest_object, "Import")

    # ---------------------------------------------------- modify_package_dict --
    # Copia exacta de CKANHarvester.modify_package_dict para garantizar
    # el mismo comportamiento de transformación de campos.

    def modify_package_dict(self, package_dict, harvest_object):
        # Map from extras key → top-level schema field name
        EXTRAS_TO_FIELDS = {
            "accrualPeriodicity": "dataset_accrualPeriodicity",
            "superTheme":         "dataset_superTheme",
            "modified":           "dataset_modified",
            "Modificado":         "dataset_modified",
            "spatial":            "dataset_spatial",
            "language":           "dataset_language",
            "dataset_theme":              "dataset_theme",
            "dataset_source":             "dataset_source",
            "dataset_publisher_name": "dataset_publisher_name"
        }

        # Required schema fields with fallback defaults if not found in extras
        FIELD_DEFAULTS = {
            "dataset_accrualPeriodicity": "http://publications.europa.eu/resource/authority/frequency/CONT",
            "dataset_description":        package_dict.get("notes", ""),
            "dataset_hvdcategory":        "categoria 1",
            "dataset_issued":             package_dict.get("metadata_created", ""),
            "dataset_modified":           package_dict.get("metadata_modified", ""),
            "dataset_status":             "http://purl.org/adms/status/Completed",
            "dataset_superTheme":         '["Sin tema"]',
            "dataset_theme":              "Tema específico 1",
        }

        # Keys that must NOT remain in extras (collide with schema fields)
        SCHEMA_COLLISION_KEYS = {"modified", "Modificado", "spatial", "temporal"}

        # 1. Promote extras to top-level fields
        remaining_extras = []
        for extra in package_dict.get("extras", []):
            key   = extra.get("key", "")
            value = extra.get("value", "")
            if key in EXTRAS_TO_FIELDS:
                field_name = EXTRAS_TO_FIELDS[key]
                if not package_dict.get(field_name):
                    package_dict[field_name] = value
            elif key in SCHEMA_COLLISION_KEYS:
                pass  # drop to avoid schema collision
            else:
                remaining_extras.append(extra)

        package_dict["extras"] = remaining_extras

        # 2. Fill in any required fields still missing
        for field, default in FIELD_DEFAULTS.items():
            #TODO:¿cómo hacer en situacion de que sea un nodo con temas?
            #if not package_dict.get(field):
            package_dict[field] = default

        # 3. Field mapping and value replacement
        if not package_dict.get("dataset_description"):
            package_dict["dataset_description"] = package_dict.get("notes", "")

        ap_value = package_dict.get("dataset_accrualPeriodicity")
        package_dict["dataset_accrualPeriodicity"] = (
            self.get_field_options("accrual_periodicity").get(ap_value, ap_value)
        )

        st_value = package_dict.get("dataset_superTheme")
        st_value = eval(st_value)  # TODO: Revisar mecanismos de seguridad para evitar inyección de código
        new_themes = [
            self.get_field_options("superthemes").get(element, element)
            for element in st_value
        ]
        package_dict["dataset_superTheme"] = new_themes[0]

        return package_dict

    # ---------------------------------------------------------------- helpers --

    def _set_config(self, config_str):
        try:
            self.config = json.loads(config_str) if config_str else {}
        except ValueError:
            self.config = {}

    def _fetch_workbook(self, url):
        ua = self.config.get("user_agent", "ckanext-harvest-xlsx/1.0")
        resp = requests.get(url, headers={"User-Agent": ua}, timeout=60)
        resp.raise_for_status()
        if "html" in resp.headers.get("Content-Type", ""):
            raise ValueError(
                "La URL devolvió HTML en vez de un archivo XLSX. "
                "Verificá que la URL apunte directamente al archivo."
            )
        return openpyxl.load_workbook(
            io.BytesIO(resp.content), read_only=True, data_only=True
        )

    def _read_sheet(self, wb, sheet_name, column_map, skip=0):
        """Lee una hoja y devuelve lista de dicts mapeados via column_map."""
        if sheet_name not in wb.sheetnames:
            log.warning("Hoja '%s' no encontrada en el workbook", sheet_name)
            return []

        rows = list(wb[sheet_name].iter_rows(values_only=True))
        if not rows:
            return []

        raw_headers = rows[skip]
        headers = [str(h).strip() if h is not None else "" for h in raw_headers]

        result = []
        for row in rows[skip + 1:]:
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            raw = {
                headers[j]: (row[j] if j < len(row) else None)
                for j in range(len(headers))
            }
            mapped = {}
            for col_header, val in raw.items():
                field_name = column_map.get(col_header)
                if field_name and val is not None:
                    mapped[field_name] = str(val).strip()
            result.append(mapped)

        return result

    def _make_guid(self, ds_id, source_url, row_index, row):
        """GUID = URL_fuente/dataset_identifier, o MD5 como fallback."""
        if ds_id:
            return "%s/%s" % (source_url.rstrip("/"), ds_id)
        title = row.get("dataset_title", "")
        raw   = "%s|%s|%d" % (source_url, title, row_index)
        return hashlib.md5(raw.encode("utf-8")).hexdigest()

    #TODO: Revisar especialmente esta función
    def _build_package_dict(self, row, harvest_object):
        """
        Convierte la fila del dataset en un package_dict con la misma
        estructura que devuelve package_show de CKAN, de modo que
        modify_package_dict() opere de forma idéntica al CKANHarvester.

        Los campos que modify_package_dict promueve desde extras se colocan
        en extras con sus claves sin prefijo (accrualPeriodicity, superTheme),
        igual que los devuelve la API CKAN.
        """
        pkg = {}

        # Campos top-level nativos de CKAN
        direct_fields = {
            "dataset_title":          "title",
            "name":                   "name",
            "owner_org":              "owner_org",
            "dataset_description":    "notes",
            "dataset_issued":         "metadata_created",
            "dataset_modified":       "metadata_modified",
            "license_id":             "license_id",
            "dataset_publisher_mbox": "author_email",
        }
        for src, dst in direct_fields.items():
            val = (row.get(src) or "").strip()
            if val:
                pkg[dst] = val

        if not pkg.get("title"):
            raise ValueError("El dataset no tiene título (dct:title)")

        pkg["name"] = self._gen_new_name(pkg.get("name") or pkg["title"])

        if not pkg.get("owner_org"):
            default_org = self.config.get("default_owner_org", "")
            if default_org:
                pkg["owner_org"] = default_org

        # Keywords → tags
        keywords_raw = row.get("dataset_keywords", "")
        if keywords_raw:
            pkg["tags"] = [
                {"name": t.strip()}
                for t in keywords_raw.split(",")
                if t.strip()
            ]

        # Extras: usar claves sin prefijo para los campos que promote
        # modify_package_dict (igual que los devuelve la API CKAN remota)
        extras = [{"key": "harvest_source_url", "value": harvest_object.source.url}]

        promoted_extras = {
            "accrualPeriodicity": (row.get("dataset_accrualPeriodicity") or "").strip(),
            "superTheme":         self._format_super_theme(row.get("dataset_superTheme", "")),
        }
        for key, val in promoted_extras.items():
            if val:
                extras.append({"key": key, "value": val})

        # Resto de campos DCAT-AP que van directamente como extras
        plain_extras = [
            "dataset_status", "dataset_theme", "dataset_hvdcategory",
            "temporal_start", "temporal_end",
            "spatial_uri", "spatial_coverage",
            "dataset_source", "dataset_identifier", "dataset_publisher_name"
        ]
        for f in plain_extras:
            val = (row.get(f) or "").strip()
            if val:
                extras.append({"key": f, "value": val})

        pkg["extras"] = extras

        # Recursos (distribuciones de la pestaña Distributions)
        resources = []
        for dist in row.get("_resources", []):
            res_url = (dist.get("url") or "").strip()
            if not res_url:
                continue
            res = {
                "url":         res_url,
                "name":        dist.get("name", pkg["title"]),
                "format":      dist.get("format", ""),
                "description": dist.get("description", ""),
            }
            for dist_field in (
                "distribution_identifier", "distribution_mediaType",
                "distribution_type", "distribution_character_set",
                "distribution_scale", "distribution_projection",
                "distribution_iso19115_url", "distribution_wfs_url",
            ):
                val = (dist.get(dist_field) or "").strip()
                if val:
                    res[dist_field] = val
            resources.append(res)

        if not resources:
            resources = [{
                "url":    harvest_object.source.url,
                "format": "XLSX",
                "name":   "Archivo fuente XLSX",
            }]

        pkg["resources"] = resources

        existing_id = self._get_existing_package_id(harvest_object)
        if existing_id:
            pkg["id"] = existing_id
        else:
            ds_identifier = (row.get("id")or "").strip()
            pkg["id"] = ds_identifier if ds_identifier else str(uuid.uuid4())

        return pkg

    def _format_super_theme(self, raw_value):
        """
        Convierte el valor de superTheme del XLSX al formato de repr de lista
        que espera modify_package_dict (ej. '["AGRI"]').
        Acepta:
          - "AGRI"           → '["AGRI"]'
          - "AGRI, SOCI"     → '["AGRI", "SOCI"]'
          - '["AGRI"]'       → '["AGRI"]'  (ya está en formato correcto)
        """
        if not raw_value:
            return ""
        raw_value = raw_value.strip()
        if raw_value.startswith("["):
            return raw_value
        items = [v.strip() for v in raw_value.split(",") if v.strip()]
        return json.dumps(items)

    def _get_existing_package_id(self, harvest_object):
        """
        Devuelve el ID interno del package si el dataset ya fue harvested,
        o None si es la primera vez.
        """
        if harvest_object.package_id:
            return harvest_object.package_id

        previous = (
            HarvestObject.Session.query(HarvestObject)
            .filter(HarvestObject.guid == harvest_object.guid)
            .filter(HarvestObject.package_id != None)   # noqa: E711
            .filter(HarvestObject.id != harvest_object.id)
            .order_by(HarvestObject.gathered.desc())
            .first()
        )
        if previous and previous.package_id:
            try:
                toolkit.get_action("package_show")(
                    {"ignore_auth": True}, {"id": previous.package_id}
                )
                return previous.package_id
            except toolkit.ObjectNotFound:
                log.warning(
                    "Package %s ya no existe en CKAN, se creará uno nuevo",
                    previous.package_id,
                )

        return None

    def _save_gather_error(self, message, harvest_job):
        log.error("XLSXHarvesterClaude gather error: %s", message)
        HarvestGatherError(message=message, job=harvest_job).save()
