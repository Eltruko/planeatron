import customtkinter as ctk
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import tkinter.messagebox as msgbox
from tkinter import filedialog

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# Verde oliva (#808000) al 40% de tinta sobre blanco.
COLOR_OLIVA_40 = "CCCC99"

class AppPlaneaciones(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Planeatrón - Asistente de Planeaciones Didácticas (NEM)")
        self.geometry("1240x800")
        self.minsize(1000, 640)

        # Título Principal
        self.lbl_title = ctk.CTkLabel(
            self, 
            text="Planeatrón - Generador de Planeaciones", 
            font=ctk.CTkFont(size=20, weight="bold")
        )
        self.lbl_title.pack(pady=(15, 5))

        # Contenedor principal de dos columnas
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=15, pady=(0, 10))
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=4, uniform="columnas")
        self.main_container.grid_columnconfigure(1, weight=5, uniform="columnas")

        # Columna izquierda: formulario con scroll (incluye el botón al final)
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_container, width=520)
        self.scroll_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=5)

        # Columna derecha: vista previa en tiempo real
        self.preview_col = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.preview_col.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=5)
        self.preview_col.grid_rowconfigure(1, weight=1)
        self.preview_col.grid_columnconfigure(0, weight=1)

        self.lbl_preview = ctk.CTkLabel(
            self.preview_col,
            text="👁 Vista previa en tiempo real",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        self.lbl_preview.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        self.preview_scroll = ctk.CTkScrollableFrame(
            self.preview_col,
            fg_color="#DCE1E5",
            corner_radius=6
        )
        self.preview_scroll.grid(row=1, column=0, sticky="nsew")

        # Hoja simulada tipo página de Word
        self.preview_page = ctk.CTkFrame(
            self.preview_scroll,
            fg_color="#FFFFFF",
            corner_radius=2,
            border_width=1,
            border_color="#B9C0C4"
        )
        self.preview_page.pack(fill="x", padx=10, pady=10)

        self.lbl_preview_nota = ctk.CTkLabel(
            self.preview_col,
            text="La vista previa es una aproximación del formato final del archivo .docx.",
            font=ctk.CTkFont(size=10),
            text_color="#5A646B",
            anchor="w"
        )
        self.lbl_preview_nota.grid(row=2, column=0, sticky="ew", pady=(5, 0))

        # --- SECCIÓN 1: DATOS GENERALES ---
        self.add_section_header("1. Datos Generales")
        self.txt_titulo = self.add_input_field(
            "Título de la Planeación:",
            placeholder="Ej. Planeación Didáctica: La noticia"
        )
        self.txt_docente = self.add_input_field("Nombre del Docente:")
        self.txt_escuela = self.add_input_field("Escuela / Institución:")
        self.txt_grado = self.add_input_field("Grado y Grupo:")
        self.txt_periodo = self.add_input_field("Periodo / Fecha:")

        # --- SECCIÓN 2: ENCUADRE CURRICULAR (NEM) ---
        self.add_section_header("2. Encuadre Curricular (NEM)")
        self.txt_campo_formativo = self.add_option_field(
            "Campo Formativo:",
            [
                "Lenguajes",
                "Saberes y Pensamiento Científico",
                "Ética, Naturaleza y Sociedad",
                "De lo Humano y lo Comunitario",
            ]
        )
        self.txt_escenario = self.add_segmented_field(
            "Escenario:",
            ["Escuela", "Aula", "Comunitario"]
        )
        self.txt_libro = self.add_combo_field(
            "Libro de Texto:",
            [
                "Múltiples Lenguajes",
                "Nuestros Saberes",
                "Proyectos de Aula",
                "Proyectos Comunitarios",
                "Proyectos Escolares",
                "Un libro sin recetas"
            ]
        )
        self.txt_problematica = self.add_textbox_field("Problemática:")
        self.txt_contenido = self.add_textbox_field("Contenido:")
        self.txt_pda = self.add_textbox_field("PDA (Proceso de Desarrollo de Aprendizaje):")
        self.txt_tema = self.add_input_field("Tema General:")
        self.txt_aprendizaje = self.add_textbox_field("Aprendizaje Esperado / Propósito:")

        # --- SECCIÓN 3: SECUENCIA DIDÁCTICA ---
        self.add_section_header("3. Secuencia Didáctica")
        self.txt_inicio = self.add_textbox_field("Inicio (Enganche / Saberes previos):")
        self.txt_desarrollo = self.add_textbox_field("Desarrollo (Actividades principales):")
        self.txt_cierre = self.add_textbox_field("Cierre (Evaluación / Puesta en común):")

        # --- SECCIÓN 4: RECURSOS Y EVALUACIÓN ---
        self.add_section_header("4. Recursos y Evaluación")
        self.txt_recursos = self.add_textbox_field("Recursos y Materiales:")
        self.txt_evaluacion = self.add_textbox_field("Criterios e Instrumentos de Evaluación:")

        # BOTÓN GENERAR (DENTRO DEL SCROLLFRAME PARA PANTALLAS PEQUEÑAS)
        self.btn_generar = ctk.CTkButton(
            self.scroll_frame, 
            text="📄 Guardar Planeación Word (.docx)", 
            command=self.generar_word,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            fg_color="#1E4D2B", 
            hover_color="#14341D"
        )
        self.btn_generar.pack(fill="x", padx=10, pady=(25, 20))

        # ESTRUCTURA FIJA Y DICCIONARIO DE ETIQUETAS PARA LA VISTA PREVIA (SIN PARPADEO)
        self._pv_labels = {}
        self._construir_estrucutra_preview()

        self._preview_snapshot = None
        self.actualizar_preview()
        self._preview_job = self.after(300, self._auto_actualizar_preview)

    def add_section_header(self, text):
        lbl = ctk.CTkLabel(
            self.scroll_frame, 
            text=text, 
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w"
        )
        lbl.pack(fill="x", pady=(15, 5))

    def add_input_field(self, label_text, placeholder=""):
        lbl = ctk.CTkLabel(self.scroll_frame, text=label_text, anchor="w")
        lbl.pack(fill="x", pady=(2, 0))
        entry = ctk.CTkEntry(self.scroll_frame, placeholder_text=placeholder)
        entry.pack(fill="x", pady=(0, 5))
        return entry

    def add_textbox_field(self, label_text):
        lbl = ctk.CTkLabel(self.scroll_frame, text=label_text, anchor="w")
        lbl.pack(fill="x", pady=(2, 0))
        tb = ctk.CTkTextbox(self.scroll_frame, height=65)
        tb.pack(fill="x", pady=(0, 5))
        return tb

    def add_option_field(self, label_text, values):
        lbl = ctk.CTkLabel(self.scroll_frame, text=label_text, anchor="w")
        lbl.pack(fill="x", pady=(2, 0))
        option = ctk.CTkOptionMenu(self.scroll_frame, values=values)
        option.pack(fill="x", pady=(0, 5))
        option.set(values[0])
        return option

    def add_segmented_field(self, label_text, values):
        lbl = ctk.CTkLabel(self.scroll_frame, text=label_text, anchor="w")
        lbl.pack(fill="x", pady=(2, 0))
        segmented = ctk.CTkSegmentedButton(self.scroll_frame, values=values)
        segmented.pack(fill="x", pady=(0, 5))
        segmented.set(values[0])
        return segmented

    def add_combo_field(self, label_text, values):
        lbl = ctk.CTkLabel(self.scroll_frame, text=label_text, anchor="w")
        lbl.pack(fill="x", pady=(2, 0))
        combo = ctk.CTkComboBox(self.scroll_frame, values=values)
        combo.pack(fill="x", pady=(0, 5))
        combo.set(values[0])
        return combo

    def _leer_campos(self, vacio="N/A"):
        def simple(widget):
            return widget.get().strip() or vacio

        def multilinea(widget):
            return widget.get("1.0", "end-1c").strip() or vacio

        return {
            "titulo": simple(self.txt_titulo),
            "docente": simple(self.txt_docente),
            "escuela": simple(self.txt_escuela),
            "grado": simple(self.txt_grado),
            "periodo": simple(self.txt_periodo),
            "campo_formativo": simple(self.txt_campo_formativo),
            "escenario": simple(self.txt_escenario),
            "libro": simple(self.txt_libro),
            "problematica": multilinea(self.txt_problematica),
            "contenido": multilinea(self.txt_contenido),
            "pda": multilinea(self.txt_pda),
            "tema": simple(self.txt_tema),
            "aprendizaje": multilinea(self.txt_aprendizaje),
            "inicio": multilinea(self.txt_inicio),
            "desarrollo": multilinea(self.txt_desarrollo),
            "cierre": multilinea(self.txt_cierre),
            "recursos": multilinea(self.txt_recursos),
            "evaluacion": multilinea(self.txt_evaluacion),
        }

    # ------------------------------------------------------------------
    # VISTA PREVIA FLUIDA (SIN DESTRUIR WIDGETS -> SIN PARPADEO)
    # ------------------------------------------------------------------
    def _construir_estrucutra_preview(self):
        margen = ctk.CTkFrame(self.preview_page, fg_color="transparent")
        margen.pack(fill="both", expand=True, padx=16, pady=14)
        margen.grid_columnconfigure(0, weight=1)

        # Título
        self._pv_labels["titulo"] = ctk.CTkLabel(
            margen, text="", font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#000000", wraplength=350, justify="center"
        )
        self._pv_labels["titulo"].grid(row=0, column=0, sticky="ew", pady=(0, 12))

        # Helper para crear tablas visuales fijas
        def crear_tabla(row_idx, filas_info, pesos=(1, 1), encabezado=None):
            marco = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
            marco.grid(row=row_idx, column=0, sticky="ew", pady=(0, 10))
            for c, w in enumerate(pesos):
                marco.grid_columnconfigure(c, weight=w)

            f_idx = 0
            if encabezado:
                ctk.CTkLabel(
                    marco, text=encabezado, fg_color="#FFFFFF", text_color="#000000",
                    font=ctk.CTkFont(size=10, weight="bold"), anchor="w", corner_radius=0
                ).grid(row=0, column=0, columnspan=len(pesos), sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
                f_idx = 1

            for r, (clave, es_bold) in enumerate(filas_info):
                lbl = ctk.CTkLabel(
                    marco, text="", fg_color="#FFFFFF", text_color="#000000",
                    font=ctk.CTkFont(size=9, weight="bold" if es_bold else "normal"),
                    anchor="w", justify="left", wraplength=260, corner_radius=0
                )
                lbl.grid(row=f_idx + r[0] if isinstance(r, tuple) else f_idx + r, 
                         column=r[1] if isinstance(r, tuple) else 1, 
                         sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
                self._pv_labels[clave] = lbl

        # 1. Tabla Datos Generales
        m_gen = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
        m_gen.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        m_gen.grid_columnconfigure(0, weight=1)
        m_gen.grid_columnconfigure(1, weight=1)

        coords = [
            ("gen_docente", 0, 0), ("gen_escuela", 0, 1),
            ("gen_grado", 1, 0), ("gen_periodo", 1, 1),
            ("gen_tema", 2, 0)
        ]
        for key, r, c in coords:
            lbl = ctk.CTkLabel(
                m_gen, text="", fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9), anchor="w", justify="left", wraplength=180, corner_radius=0
            )
            lbl.grid(row=r, column=c, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
            self._pv_labels[key] = lbl

        # 2. Encuadre NEM
        m_nem = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
        m_nem.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        m_nem.grid_columnconfigure(0, weight=1)
        m_nem.grid_columnconfigure(1, weight=2)
        ctk.CTkLabel(
            m_nem, text="ENCUADRE CURRICULAR (NEM)", fg_color="#FFFFFF", text_color="#000000",
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w", corner_radius=0
        ).grid(row=0, column=0, columnspan=2, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)

        items_nem = ["campo_formativo", "escenario", "libro", "problematica", "contenido", "pda"]
        titulos_nem = ["Campo Formativo", "Escenario", "Libro de Texto", "Problemática", "Contenido", "PDA"]
        
        for idx, (key, tit) in enumerate(zip(items_nem, titulos_nem), start=1):
            ctk.CTkLabel(
                m_nem, text=tit, fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9, weight="bold"), anchor="w", corner_radius=0
            ).grid(row=idx, column=0, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)

            lbl = ctk.CTkLabel(
                m_nem, text="", fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9), anchor="w", justify="left", wraplength=240, corner_radius=0
            )
            lbl.grid(row=idx, column=1, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
            self._pv_labels[f"nem_{key}"] = lbl

        # 3. Aprendizaje Esperado
        m_app = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
        m_app.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        m_app.grid_columnconfigure(0, weight=1)
        lbl_app = ctk.CTkLabel(
            m_app, text="", fg_color="#FFFFFF", text_color="#000000",
            font=ctk.CTkFont(size=9), anchor="w", justify="left", wraplength=360, corner_radius=0
        )
        lbl_app.grid(row=0, column=0, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
        self._pv_labels["aprendizaje"] = lbl_app

        # 4. Secuencia Didáctica
        m_seq = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
        m_seq.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        m_seq.grid_columnconfigure(0, weight=1)
        m_seq.grid_columnconfigure(1, weight=2)
        ctk.CTkLabel(
            m_seq, text="Fase de la Clase  |  Actividades", fg_color="#FFFFFF", text_color="#000000",
            font=ctk.CTkFont(size=10, weight="bold"), anchor="w", corner_radius=0
        ).grid(row=0, column=0, columnspan=2, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)

        fases = [("Inicio", "seq_inicio"), ("Desarrollo", "seq_desarrollo"), ("Cierre", "seq_cierre")]
        for idx, (fase_nom, key) in enumerate(fases, start=1):
            ctk.CTkLabel(
                m_seq, text=fase_nom, fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9, weight="bold"), anchor="w", corner_radius=0
            ).grid(row=idx, column=0, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)

            lbl = ctk.CTkLabel(
                m_seq, text="", fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9), anchor="w", justify="left", wraplength=240, corner_radius=0
            )
            lbl.grid(row=idx, column=1, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
            self._pv_labels[key] = lbl

        # 5. Recursos y Evaluación
        m_rec = ctk.CTkFrame(margen, fg_color="#B9C0C4", corner_radius=0)
        m_rec.grid(row=5, column=0, sticky="ew", pady=(0, 10))
        m_rec.grid_columnconfigure(0, weight=1)
        m_rec.grid_columnconfigure(1, weight=2)

        rec_items = [("Recursos y Materiales:", "rec_recursos"), ("Evaluación:", "rec_evaluacion")]
        for idx, (tit, key) in enumerate(rec_items):
            ctk.CTkLabel(
                m_rec, text=tit, fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9, weight="bold"), anchor="w", corner_radius=0
            ).grid(row=idx, column=0, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)

            lbl = ctk.CTkLabel(
                m_rec, text="", fg_color="#FFFFFF", text_color="#000000",
                font=ctk.CTkFont(size=9), anchor="w", justify="left", wraplength=240, corner_radius=0
            )
            lbl.grid(row=idx, column=1, sticky="nsew", padx=1, pady=1, ipadx=5, ipady=3)
            self._pv_labels[key] = lbl

    def _auto_actualizar_preview(self):
        if not self.winfo_exists():
            return
        datos = self._leer_campos(vacio="")
        if datos != self._preview_snapshot:
            self.actualizar_preview()
        self._preview_job = self.after(300, self._auto_actualizar_preview)

    def actualizar_preview(self):
        if not self.winfo_exists():
            return

        datos = self._leer_campos(vacio="")
        self._preview_snapshot = datos

        def v(val):
            return val if val else "(Sin capturar)"

        # Actualizamos únicamente el texto de cada widget existente
        self._pv_labels["titulo"].configure(text=datos["titulo"] or "PLANEACIÓN DIDÁCTICA")
        self._pv_labels["gen_docente"].configure(text=f"Docente: {v(datos['docente'])}")
        self._pv_labels["gen_escuela"].configure(text=f"Escuela: {v(datos['escuela'])}")
        self._pv_labels["gen_grado"].configure(text=f"Grado/Grupo: {v(datos['grado'])}")
        self._pv_labels["gen_periodo"].configure(text=f"Periodo: {v(datos['periodo'])}")
        self._pv_labels["gen_tema"].configure(text=f"Tema: {v(datos['tema'])}")

        self._pv_labels["nem_campo_formativo"].configure(text=v(datos["campo_formativo"]))
        self._pv_labels["nem_escenario"].configure(text=v(datos["escenario"]))
        self._pv_labels["nem_libro"].configure(text=v(datos["libro"]))
        self._pv_labels["nem_problematica"].configure(text=v(datos["problematica"]))
        self._pv_labels["nem_contenido"].configure(text=v(datos["contenido"]))
        self._pv_labels["nem_pda"].configure(text=v(datos["pda"]))

        self._pv_labels["aprendizaje"].configure(text=f"APRENDIZAJE ESPERADO / PROPÓSITO:\n{v(datos['aprendizaje'])}")

        self._pv_labels["seq_inicio"].configure(text=v(datos["inicio"]))
        self._pv_labels["seq_desarrollo"].configure(text=v(datos["desarrollo"]))
        self._pv_labels["seq_cierre"].configure(text=v(datos["cierre"]))

        self._pv_labels["rec_recursos"].configure(text=v(datos["recursos"]))
        self._pv_labels["rec_evaluacion"].configure(text=v(datos["evaluacion"]))

    # ------------------------------------------------------------------
    # GENERACIÓN DEL DOCUMENTO WORD (.DOCX)
    # ------------------------------------------------------------------
    def set_cell_background(self, cell, fill_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def escribir_celda_titulo(self, cell, texto):
        cell.text = texto
        self.set_cell_background(cell, COLOR_OLIVA_40)
        for parrafo in cell.paragraphs:
            for run in parrafo.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)

    def escribir_dato(self, cell, etiqueta, valor):
        parrafo = cell.paragraphs[0]
        run_etiqueta = parrafo.add_run(f"{etiqueta}: ")
        run_etiqueta.font.bold = True
        parrafo.add_run(valor)

    def aplicar_texto_negro(self, doc):
        def pintar(parrafos):
            for parrafo in parrafos:
                for run in parrafo.runs:
                    run.font.color.rgb = RGBColor(0, 0, 0)

        pintar(doc.paragraphs)
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    pintar(celda.paragraphs)

    def generar_word(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Documento de Word", "*.docx"), ("Todos los archivos", "*.*")],
            title="Guardar Planeación Didáctica",
            initialfile="Planeacion_Didactica.docx"
        )
        
        if not filepath:
            return

        datos = self._leer_campos()

        doc = Document()

        title = doc.add_heading(datos["titulo"], level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Datos Generales
        table_gen = doc.add_table(rows=3, cols=2, style="Table Grid")
        table_gen.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        data_gen = [
            ("Docente", datos["docente"], "Escuela", datos["escuela"]),
            ("Grado/Grupo", datos["grado"], "Periodo", datos["periodo"]),
            ("Tema General", datos["tema"], "", ""),
        ]

        for i, (e_izq, v_izq, e_der, v_der) in enumerate(data_gen):
            celdas = table_gen.rows[i].cells
            self.escribir_dato(celdas[0], e_izq, v_izq)
            if e_der:
                self.escribir_dato(celdas[1], e_der, v_der)

        doc.add_paragraph()

        # Encuadre NEM
        table_nem = doc.add_table(rows=7, cols=2, style="Table Grid")
        table_nem.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_nem = table_nem.cell(0, 0).merge(table_nem.cell(0, 1))
        self.escribir_celda_titulo(hdr_nem, "ENCUADRE CURRICULAR (NEM)")

        data_nem = [
            ("Campo Formativo", datos["campo_formativo"]),
            ("Escenario", datos["escenario"]),
            ("Libro de Texto", datos["libro"]),
            ("Problemática", datos["problematica"]),
            ("Contenido", datos["contenido"]),
            ("PDA", datos["pda"]),
        ]

        for idx, (etiqueta, valor) in enumerate(data_nem, start=1):
            fila_nem = table_nem.rows[idx].cells
            self.escribir_celda_titulo(fila_nem[0], etiqueta)
            fila_nem[1].text = valor

        doc.add_paragraph()

        # Aprendizaje Esperado
        table_app = doc.add_table(rows=2, cols=1, style="Table Grid")
        table_app.alignment = WD_TABLE_ALIGNMENT.CENTER
        self.escribir_celda_titulo(table_app.cell(0, 0), "APRENDIZAJE ESPERADO / PROPÓSITO:")
        table_app.cell(1, 0).text = datos["aprendizaje"]

        doc.add_paragraph()

        # Secuencia Didáctica
        table_seq = doc.add_table(rows=4, cols=2, style="Table Grid")
        table_seq.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_cells = table_seq.rows[0].cells
        self.escribir_celda_titulo(hdr_cells[0], "Fase de la Clase")
        self.escribir_celda_titulo(hdr_cells[1], "Actividades")

        fases = [
            ("Inicio", datos["inicio"]),
            ("Desarrollo", datos["desarrollo"]),
            ("Cierre", datos["cierre"])
        ]

        for idx, (fase_nombre, fase_texto) in enumerate(fases, start=1):
            row_cells = table_seq.rows[idx].cells
            row_cells[0].text = fase_nombre
            self.set_cell_background(row_cells[0], COLOR_OLIVA_40)
            row_cells[1].text = fase_texto

        doc.add_paragraph()

        # Recursos y Evaluación
        table_rec = doc.add_table(rows=2, cols=2, style="Table Grid")

        r1 = table_rec.rows[0].cells
        self.escribir_celda_titulo(r1[0], "Recursos y Materiales:")
        r1[1].text = datos["recursos"]

        r2 = table_rec.rows[1].cells
        self.escribir_celda_titulo(r2[0], "Evaluación:")
        r2[1].text = datos["evaluacion"]

        self.aplicar_texto_negro(doc)

        try:
            doc.save(filepath)
            msgbox.showinfo("Éxito", f"¡Planeación guardada correctamente en:\n{filepath}")
        except Exception as e:
            msgbox.showerror("Error", f"No se pudo guardar el archivo.\nDetalle: {e}")

if __name__ == "__main__":
    app = AppPlaneaciones()
    app.mainloop()
