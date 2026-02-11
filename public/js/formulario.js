document.addEventListener("DOMContentLoaded", function () {
  console.log("El DOM se ha cargado");
  const sexocombo = document.getElementById("sexo-opciones");
  const sexoconyugecombo = document.getElementById("sexoConyuge-opciones");
  const civilcombo = document.getElementById("civil-opciones");
  const sangrecombo = document.getElementById("sangre-opciones");
  const botonFormulario = document.getElementById("botonFormulario");

  const camposAcentos = document.querySelectorAll(".acentos");
  const inputsFecha = document.querySelectorAll(
    "input[data-format='dd-mm-yyyy']",
  );

  const radioscronica = document.getElementsByName("cronica");
  const hijoscontador = document.getElementById("hijoscontador");

  const radiosEmbarazo = document.getElementsByName("embarazo");
  const preguntaEmbarazo = document.getElementById("pregunta-embarazo");
  const preguntaHijos = document.getElementById("preguntahijos");

  var radiospapa = document.getElementsByName("papa");
  var form = document.querySelector("form");

  const fechaConyuge = document.getElementById("fechaConyuge");

  const telefonoInput = document.getElementById("noTelefono");

  function actualizarPreguntaEmbarazo() {
    const sexo = sexocombo.value;

    if (sexo == "Femenino") {
      preguntaEmbarazo.classList.remove("hidden");
      // Requiere al menos uno seleccionado
      radiosEmbarazo.forEach((radio) => (radio.required = true));
      console.log("se ha agregado el required");
    } else {
      preguntaEmbarazo.classList.add("hidden");
      radiosEmbarazo.forEach((radio) => (radio.required = false));

      // Opcional: deseleccionar
      // radiosEmbarazo.forEach(radio => radio.checked = false);

      console.log("required removido");
    }
  }

  function actualizarPreguntaConyuge() {
    const estadoCivil = document.getElementById("civil-opciones").value;
    const datosConyuge = document.getElementById("datosConyuge");
    const camposConyuge = document.getElementsByClassName("conyuge");

    console.log(estadoCivil);

    if (estadoCivil === "Casado/a" || estadoCivil === "Unionlibre") {
      datosConyuge.classList.remove("hidden");
      for (let i = 0; i < camposConyuge.length; i++) {
        camposConyuge[i].required = true;
      }
    } else {
      datosConyuge.classList.add("hidden");
      for (let i = 0; i < camposConyuge.length; i++) {
        camposConyuge[i].required = false;
      }
    }
  }

  function actualizarPreguntaHijos() {
    if (radiospapa[0].checked) {
      preguntaHijos.classList.remove("hidden");
      hijoscontador.setAttribute("required", "true");
    } else {
      const espacioHijos = document.getElementById("espacioHijos");
      preguntaHijos.classList.add("hidden");
      hijoscontador.removeAttribute("required");
      while (espacioHijos.hasChildNodes()) {
        espacioHijos.removeChild(espacioHijos.lastChild);
      }
      hijoscontador.value = null;
    }
  }
  function selectEventListener() {
    const selectsDinamicos = document.querySelectorAll(
      'select[name^="selectHijo"]',
    );
    for (let o = 0; o < selectsDinamicos.length; o++) {
      selectsDinamicos[o].addEventListener("change", function () {
        if (selectsDinamicos[o].value !== "Seleccione") {
          selectsDinamicos[o].setCustomValidity("");
        } else {
          selectsDinamicos[o].setCustomValidity("Seleccione una opción");
        }
        selectsDinamicos[o].reportValidity();
      });
    }
  }

  /**
   * Configura la visibilidad de un contenedor basada en la selección de un grupo de radio buttons.
   * @param {string} radioName - El atributo 'name' del grupo de radio buttons.
   * @param {string} containerId - El 'id' del contenedor a mostrar/ocultar.
   * @param {string} showValue - El 'value' del radio button que debe mostrar el contenedor.
   */
  function configurarVisibilidadCondicional(radioName, containerId, showValue) {
    const radios = document.querySelectorAll(`input[name="${radioName}"]`);
    const container = document.getElementById(containerId);

    // Salir si los elementos no existen para evitar errores.
    if (!radios.length || !container) {
      return;
    }

    const requiredElements = container.querySelectorAll("input, select");

    const manejarVisibilidad = () => {
      // Encontrar el radio que está actualmente seleccionado.
      const radioSeleccionado = document.querySelector(
        `input[name="${radioName}"]:checked`,
      );

      // Si por alguna razón ninguno está seleccionado, no hacer nada.
      if (!radioSeleccionado) return;

      if (radioSeleccionado.value === showValue) {
        container.classList.remove("hidden");
        requiredElements.forEach((el) => el.setAttribute("required", true));
      } else {
        container.classList.add("hidden");
        requiredElements.forEach((el) => {
          el.removeAttribute("required");
          // Limpiar valor para diferentes tipos de input.
          if (el.type === "checkbox" || el.type === "radio") {
            el.checked = false;
          } else {
            el.value = "";
          }
        });
      }
    };

    // Añadir el listener a cada radio del grupo.
    radios.forEach((radio) => {
      radio.addEventListener("change", manejarVisibilidad);
    });

    // Ejecutar la función una vez al cargar la página para establecer el estado inicial.
    manejarVisibilidad();
  }

  function agregarCamposHijos() {
    const noHijos = parseInt(hijoscontador.value) || 0;
    const espacioHijos = document.getElementById("espacioHijos");
    const hijosActuales = espacioHijos.children.length;

    // Si hay menos hijos de los deseados, eliminamos los sobrantes del final
    while (espacioHijos.children.length > noHijos) {
      espacioHijos.removeChild(espacioHijos.lastChild);
    }

    // Si hay más hijos deseados que actuales, agregamos los que faltan
    for (let i = hijosActuales; i < noHijos; i++) {
      const divHijo = document.createElement("div");
      divHijo.classList.add(
        "mb-3",
        "p-4",
        "rounded",
        "border",
        "border-gray-300",
        "hijo-item",
      );

      // Título del hijo
      const titulo = document.createElement("h3");
      titulo.textContent = `Datos del Hijo ${i + 1}`;
      titulo.classList.add("font-bold", "text-lg", "mb-4", "text-gray-800");
      divHijo.appendChild(titulo);

      // Iniciales label
      const labelIniciales = document.createElement("label");
      labelIniciales.classList.add(
        "block",
        "font-medium",
        "mb-1",
        "text-gray-700",
      );

      const spanRojo = document.createElement("span");
      spanRojo.textContent = "Nombre completo";
      spanRojo.style.color = "red";

      labelIniciales.appendChild(spanRojo);

      labelIniciales.append(
        " empezando por apellidos de la hija o hijo (sin acentos):",
      );

      divHijo.appendChild(labelIniciales);

      const input = document.createElement("input");
      input.type = "text";
      input.name = "inicialesHijo" + i;
      input.placeholder = "";
      input.classList.add(
        "border",
        "rounded-full",
        "bg-gray-50",
        "m-2",
        "p-3",
        "uppercase",
      );
      input.setAttribute("required", "true");
      divHijo.appendChild(input);

      // Sexo label
      const labelSexo = document.createElement("label");
      labelSexo.textContent = "Género de la hija o hijo:";
      labelSexo.classList.add(
        "block",
        "font-medium",
        "mb-1",
        "text-gray-700",
        "mt-4",
      );
      divHijo.appendChild(labelSexo);

      // Sexo select
      const select = document.createElement("select");
      select.name = "selectHijo" + i;
      select.classList.add(
        "border",
        "rounded-full",
        "bg-gray-50",
        "m-2",
        "p-3",
      );
      select.setAttribute("required", "true");

      const opcion1 = document.createElement("option");
      opcion1.value = "Seleccione";
      opcion1.textContent = "Seleccione el género";
      select.appendChild(opcion1);

      const opcion2 = document.createElement("option");
      opcion2.value = "Masculino";
      opcion2.textContent = "Masculino";
      select.appendChild(opcion2);

      const opcion3 = document.createElement("option");
      opcion3.value = "Femenino";
      opcion3.textContent = "Femenino";
      select.appendChild(opcion3);

      select.setCustomValidity("Seleccione el género del hijo");

      select.addEventListener("change", function () {
        if (select.value !== "Seleccione") {
          select.setCustomValidity("");
        } else {
          select.setCustomValidity("Seleccione el género del hijo");
        }
        select.reportValidity();
      });
      divHijo.appendChild(select);

      // Fecha nacimiento label
      const labelFecha = document.createElement("label");
      labelFecha.textContent = "Fecha de nacimiento de la hija o hijo:";
      labelFecha.classList.add(
        "block",
        "font-medium",
        "mb-1",
        "text-gray-700",
        "mt-4",
      );
      divHijo.appendChild(labelFecha);

      // Fecha nacimiento input
      const input2 = document.createElement("input");
      input2.type = "text";
      input2.name = "fechaNacimientoHijo" + i;
      input2.placeholder = "dd-mm-yyyy";
      input2.classList.add(
        "border",
        "rounded-full",
        "bg-gray-50",
        "m-2",
        "p-3",
        "datepicker",
      );
      input2.setAttribute("required", "true");
      Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(input2);
      divHijo.appendChild(input2);

      // Edad label
      //const labelEdad = document.createElement("label");
      //labelEdad.textContent = "Edad actual de la hija o hijo en años (Ejemplo: 5):";
      //labelEdad.classList.add("block", "font-medium", "mb-1", "text-gray-700", "mt-4");
      //divHijo.appendChild(labelEdad);

      // Edad input
      //const input3 = document.createElement("input");
      //input3.type = "number";
      //input3.name = "edadHijo" + i;
      //input3.placeholder = "Ejemplo: 5";
      //input3.classList.add("border", "rounded-full", "bg-gray-50", "m-2", "p-3");
      //input3.setAttribute("required", "true");
      //divHijo.appendChild(input3);

      espacioHijos.appendChild(divHijo);

      // Flatpickr para fechas
      flatpickr(input2, {
        dateFormat: "d-m-Y",
        allowInput: true,
        maxDate: "today",
      });

      // Listener select
      selectEventListener();
    }
  }

  sexocombo.addEventListener("change", function () {
    if (sexocombo.value !== "Seleccione") {
      sexocombo.setCustomValidity("");
    } else {
      sexocombo.setCustomValidity(
        "Rellene correctamente: seleccione una opción",
      );
    }
  });

  civilcombo.addEventListener("change", function () {
    if (civilcombo.value !== "Seleccione") {
      civilcombo.setCustomValidity("");
    } else {
      civilcombo.setCustomValidity(
        "Rellene correctamente: seleccione un estado civil",
      );
    }

    if (civilcombo.value === "Casado/a") {
      sexoconyugecombo.addEventListener("change", function () {
        if (sexoconyugecombo.value !== "Seleccione") {
          sexoconyugecombo.setCustomValidity("");
        } else {
          sexoconyugecombo.setCustomValidity(
            "Rellene correctamente: seleccione una opción",
          );
        }
      });
    }
  });

  sangrecombo.addEventListener("change", function () {
    if (sangrecombo.value !== "Seleccione") {
      sangrecombo.setCustomValidity("");
    } else {
      sangrecombo.setCustomValidity(
        "Rellene correctamente: seleccione una opción",
      );
    }
  });
  form.addEventListener("submit", function (event) {
    let valido = true;
    const telefonoInput = document.getElementById("noTelefono");
    const postalInput = document.getElementById("postal");
    const errorMessage = document.getElementById("error-message");
    const errorMessage2 = document.getElementById("error-message2");

    // Resetear mensajes de error
    errorMessage.style.display = "none";
    errorMessage2.style.display = "none";

    for (let i = 1; i <= 3; i++) {
      const titulo = document.querySelector(`[name="licenciatura_titulo${i}"]`);
      const fecha = document.querySelector(`[name="licenciatura_fecha${i}"]`);
      const inst = document.querySelector(
        `[name="licenciatura_institucion${i}"]`,
      );

      if (!titulo || !fecha || !inst) continue;

      const hayAlgunDato =
        titulo.value.trim() || fecha.value.trim() || inst.value.trim();

      if (
        hayAlgunDato &&
        (!titulo.value.trim() || !fecha.value.trim() || !inst.value.trim())
      ) {
        alert(
          `⚠️ Completa todos los campos de la Licenciatura ${i} o deja todos vacíos.`,
        );
        if (!titulo.value.trim()) titulo.focus();
        else if (!fecha.value.trim()) fecha.focus();
        else inst.focus();
        e.preventDefault();
        valido = false;
        break;
      }
    }
    // Validación de sexo
    if (sexocombo.value === "Seleccione") {
      sexocombo.setCustomValidity(
        "Rellene correctamente: seleccione una opción",
      );
      sexocombo.reportValidity();
      valido = false;
    } else {
      sexocombo.setCustomValidity("");
    }

    if (civilcombo.value === "Seleccione") {
      civilcombo.setCustomValidity(
        "Rellene correctamente: seleccione una opción",
      );
      civilcombo.reportValidity();
      valido = false;
    } else if (civilcombo.value === "Casado/a") {
      civilcombo.setCustomValidity("");

      if (sexoconyugecombo.value === "Seleccione") {
        sexoconyugecombo.setCustomValidity(
          "Rellene correctamente: seleccione una opción",
        );
        sexoconyugecombo.reportValidity();
        valido = false;
      } else {
        sexoconyugecombo.setCustomValidity("");
      }
    } else {
      civilcombo.setCustomValidity("");
      sexoconyugecombo.setCustomValidity("");
    }

    if (sangrecombo.value === "Seleccione") {
      sangrecombo.setCustomValidity(
        "Rellene correctamente: seleccione una opción",
      );
      sangrecombo.reportValidity();
      valido = false;
    } else {
      sangrecombo.setCustomValidity("");
    }

    // Validación de teléfono (corregido isNaN)
    const telefono = telefonoInput.value;
    if (telefono.length !== 10 || isNaN(telefono)) {
      errorMessage.style.display = "inline";
      valido = false;
    }

    // Validación de código postal (mejor usar input text)
    const codigoPostal = postalInput.value;
    if (codigoPostal.length !== 5 || isNaN(codigoPostal)) {
      errorMessage2.style.display = "inline";
      valido = false;
    }

    if (!validarEscolaridad("licenciatura", 3, "Licenciatura")) valido = false;
    if (!validarEscolaridad("maestria", 3, "Maestría")) valido = false;
    if (!validarEscolaridad("especialidad", 3, "Especialidad")) valido = false;
    if (!validarEscolaridad("doctorado", 2, "Doctorado")) valido = false;

    if (!validarActualizaciones(10)) valido = false;
    // Si alguna validación falla, prevenir envío
    if (!valido) {
      event.preventDefault();
      // Opcional: enfocar primer campo con error
      if (sexocombo.value === "Seleccione") sexocombo.focus();
      else if (civilcombo.value === "Seleccione") civilcombo.focus();
      else if (
        civilcombo.value === "Casado/a" &&
        sexoconyugecombo.value === "Seleccione"
      )
        sexoconyugecombo.focus();
      else if (sangrecombo.value === "Seleccione") civilcombo.focus();
      else if (errorMessage.style.display === "inline") telefonoInput.focus();
      else postalInput.focus();
    }

    if (valido) {
      event.preventDefault();

      const aceptado = confirm(`AVISO DE PRIVACIDAD
              Plataforma de Actualización de Datos Personales del Personal del Poder Judicial del Estado de Hidalgo

¿Quién es el responsable del tratamiento de sus datos personales? 
El Poder Judicial del Estado de Hidalgo, a través de la Dirección de Recursos Humanos del Consejo de la Judicatura ubicada en el kilómetro 84.5, Carretera México Pachuca, Sector Primario, C.P. 42080 Pachuca de Soto, Hidalgo. 

¿Qué datos personales recabamos y para que finalidad? 
Los datos personales son recabados para las siguientes finalidades: 
Actualizar el expediente físico y la base de datos, de los trabajadores del Poder Judicial.
Generar reportes.
Elaborar reportes estadísticos en los que los datos serán tratados de forma disociada.
Dar trámite a solicitudes de información.


Los datos personales que se recaban para las finalidades antes señaladas son: 
1. Datos identificativos: Nombre, lugar y fecha de nacimiento, sexo, Clave Única de Registro de Población (CURP), RFC, estado civil, nombre de cónyuge, nombre, sexo, fecha de nacimiento y edad de descendientes en primer grado, acta de nacimiento, domicilios, código postal, números telefónicos personales y de casa, nombre y número telefónico de un familiar, correo electrónico;
2. Datos académicos: Grados de estudios profesionales, título profesional, cédula profesional, certificados, constancias de estudios, fecha de obtención de documento académico y actualización academica; 
3. Datos laborales: Número de empleado, nombramiento o designación de cargo, fecha de alta en el cargo, número de seguridad social, experiencia laboral, actas y procedimientos de responsabilidades y curriculum vitae; 
4. Datos patrimoniales: Registro Federal de Contribuyente (RFC), actividad económica y régimen fiscal en cédulas del RFC; 
5. Datos biométricos: grupo sanguíneo. 


¿Cuál es el fundamento legal que permite recabar dichos datos personales? 
La Dirección de Recursos Humanos está facultada para tratar los datos personales relacionados con los asuntos de su competencia de conformidad con las atribuciones que le confiere la Ley Orgánica del Poder Judicial del Estado de Hidalgo en su artículo 142, fracción I y Ley de Protección de Datos Personales en Posesión de Sujetos Obligados para el Estado de Hidalgo en sus artículos 12 y 15. 
¿Con quién se comparten los datos personales? 
La Dirección de Recursos Humanos puede realizar transferencias con el Servicio de Administración Tributaria (SAT), Instituto de Seguridad y Servicios Sociales de los Trabajadores del Estado (ISSSTE), Fondo de la Vivienda del Instituto de Seguridad y Servicios Sociales de los Trabajadores del Estado (FOVISSSTE), Auditoria Superior de la Federación (ASF), Auditoria Superior del Estado de Hidalgo (ASEH), Gobierno del Estado de Hidalgo, y aquellas que sean necesarias para cumplir con lo dispuesto en materia de transparencia, acceso a la información pública y protección de datos personales, así como para atender solicitudes o requerimientos realizados por autoridades competentes, de conformidad con los artículos 19, 98 y 99 de la Ley de Protección de Datos Personales en Posesión de Sujetos Obligados para el Estado de Hidalgo, en las que no se requerirá de su consentimiento. 


¿Dónde puedo ejercer los derechos de acceso, rectificación, cancelación y oposición? 
Los derechos de acceso, rectificación, cancelación y oposición de sus datos personales (Derechos ARCO) que se encuentran en posesión del Poder Judicial del Estado de Hidalgo, podrán ejercerse a través de los siguientes medios: 
1. Directamente en el domicilio de la Unidad de Transparencia del Poder Judicial sito en Carretera México - Pachuca km. 84.5 Centro Cívico, Sector Primario, C.P. 42080, Pachuca de Soto, Hidalgo. 
2. Por escrito dirigido a la Unidad de Transparencia del Poder Judicial 
3. Por internet a través de la Plataforma Nacional de Transparencia en el siguiente enlace: https://www.plataformadetransparencia.org.mx/group/guest/sisai_solicitudes#/datospersonales 
4. Por correo electrónico a la cuenta de la Unidad de Transparencia del Poder Judicial, en la siguiente dirección: transparencia@pjhidalgo.gob.mx


Si desea más información sobre como ejercer sus derechos ARCO, favor de comunicarse a la Unidad de Transparencia del Poder Judicial al 771 7179000 extensión 9546. 
¿Dónde puedo consultar el aviso de privacidad y por qué medios se comunican a los titulares de los datos personales los cambios realizados? 
El presente aviso de privacidad está sujeto a modificaciones, cambios o actualizaciones derivadas de causas diversas como lo pueden ser nuevos requerimientos legales, cambios de domicilio o medios de contacto, lo cual se informará a través de la página de internet www.pjhidalgo.gob.mx 
Lo anterior en cumplimiento a lo dispuesto en los artículos 67 fracción III de la Ley de Transparencia y Acceso a la Información Pública del Estado de Hidalgo; 3 fracción I, II, VII, VIII, IX, XXXV,19, 24, 34, 35, 36, 37, 39 y 42 de la Ley de Protección de Datos Personales en Posesión de Sujetos Obligados para el Estado de Hidalgo.

          ¿Deseas continuar y enviar tu información?`);

      if (aceptado) {
        botonFormulario.disabled = true;
        botonFormulario.textContent = "Enviando...";
        form.submit(); // ✅ Solo se envía si se acepta
      }
    }
  });

  sexocombo.addEventListener("change", actualizarPreguntaEmbarazo);
  civilcombo.addEventListener("change", actualizarPreguntaConyuge);
  radiospapa.forEach((radio) => {
    radio.addEventListener("change", actualizarPreguntaHijos);
  });

  hijoscontador.addEventListener("input", agregarCamposHijos);

  // Sincronización con HTML generado por EJS
  const hijosExistentes = document.querySelectorAll(".hijo-item");
  if (hijosExistentes.length > 0) {
    // 1. Actualizar el contador visual si no vino lleno
    if (!hijoscontador.value) {
      hijoscontador.value = hijosExistentes.length;
    }

    // 2. Mostrar la sección de hijos
    preguntaHijos.classList.remove("hidden");
    hijoscontador.setAttribute("required", "true");

    // 3. Inicializar listeners en los elementos ya existentes
    hijosExistentes.forEach((div) => {
      // Inicializar Flatpickr y Inputmask
      const dateInput = div.querySelector(".datepicker");
      if (dateInput) {
        flatpickr(dateInput, {
          dateFormat: "d-m-Y",
          allowInput: true,
          maxDate: "today",
          defaultDate: dateInput.value, // Toma el valor ISO del HTML
        });
        Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(dateInput);
      }

      // Inicializar validación de selects
      const select = div.querySelector("select");
      if (select) {
        select.addEventListener("change", function () {
          if (select.value !== "Seleccione") {
            select.setCustomValidity("");
          } else {
            select.setCustomValidity("Seleccione el género del hijo");
          }
          select.reportValidity();
        });
      }

      // Inicializar validación de acentos
      const textInput = div.querySelector(
        'input[type="text"]:not(.datepicker)',
      );
      if (textInput) {
        reemplazarAcentos(textInput);
      }
    });
  }

  // Llamadas a la nueva función refactorizada
  configurarVisibilidadCondicional("cronica", "preguntaCronica", "1");
  configurarVisibilidadCondicional(
    "coincideDomicilio",
    "domicilioActualContainer",
    "0",
  );
  configurarVisibilidadCondicional(
    "indigena",
    "comunidadIndigenaContainer",
    "1",
  );
  configurarVisibilidadCondicional(
    "hablaLenguaIndigena",
    "familiaLingContainer",
    "1",
  );
  configurarVisibilidadCondicional(
    "discapacidad",
    "tipoDiscapacidadContainer",
    "1",
  );
  configurarVisibilidadCondicional("lenguaSenias", "nivelLSMContainer", "Sí");
  configurarVisibilidadCondicional("tieneTecnica", "tecnicaContainer", "1");
  configurarVisibilidadCondicional(
    "tieneBachillerato",
    "bachilleratoContainer",
    "1",
  );

  // Sincronización de Estudios Superiores desde BD
  const contadoresEstudios = [
    "contadorLicenciaturasBD",
    "contadorMaestriaBD",
    "contadorDoctoradoBD",
    "contadorPosdoctoradoBD",
    "contadorEspecialidadBD",
  ];
  const hayEstudiosCargados = contadoresEstudios.some((id) => {
    const el = document.getElementById(id);
    return el && parseInt(el.value) > 0;
  });

  if (hayEstudiosCargados) {
    configurarFechaInput();
    procesarCamposAcentos();
  }

  // Sincronización de Actualizaciones Profesionales desde BD
  const contadorActBD = document.getElementById("contadorActualizacionBD");
  if (contadorActBD && parseInt(contadorActBD.value) > 0) {
    actualizacionIndex = parseInt(contadorActBD.value);
    document.getElementById("actualizacionTotal").value = actualizacionIndex;
    configurarFechaInput();
    procesarCamposAcentos();
  }

  // Sincronización de Experiencia Laboral desde BD
  const contadorExpBD = document.getElementById("contadorExperienciaBD");
  if (contadorExpBD && parseInt(contadorExpBD.value) > 0) {
    experienciaIndex = parseInt(contadorExpBD.value);
    document.getElementById("experienciaTotal").value = experienciaIndex;

    // Reactivar listeners para elementos precargados
    for (let i = 0; i < experienciaIndex; i++) {
      // Lógica Institución (PJ vs OTRA)
      const selectInst = document.querySelector(
        `select[name="experiencia_institucion${i}"]`,
      );
      const adscripContainer = document.getElementById(
        `adscripcion-container-${i}`,
      );
      const otraContainer = document.getElementById(`otra-inst-container-${i}`);

      if (selectInst) {
        selectInst.addEventListener("change", function (e) {
          if (e.target.value === "PJ") {
            adscripContainer.classList.remove("hidden");
            otraContainer.classList.add("hidden");
          } else {
            adscripContainer.classList.add("hidden");
            otraContainer.classList.remove("hidden");
          }
        });
      }

      // Lógica Checkbox "A la fecha"
      const finInput = document.getElementById(`experiencia_fin${i}`);
      const actualCb = document.getElementById(`experiencia_actual${i}`);

      if (actualCb) {
        actualCb.addEventListener("change", () => {
          if (actualCb.checked) {
            finInput.classList.add("hidden");
            finInput.removeAttribute("required");
            finInput.value = "";
          } else {
            finInput.classList.remove("hidden");
            finInput.setAttribute("required", "true");
          }
        });
      }
    }

    configurarFechaInput();
    procesarCamposAcentos();
  }

  // Ejecutar validaciones de visibilidad iniciales para secciones manuales
  actualizarPreguntaHijos();
  actualizarPreguntaEmbarazo();
  actualizarPreguntaConyuge();

  Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(fechaConyuge);
  flatpickr(fechaConyuge, {
    dateFormat: "d-m-Y",
    allowInput: true,
    maxDate: "today",
  });

  inputsFecha.forEach(function (input) {
    // Flatpickr
    flatpickr(input, {
      dateFormat: "d-m-Y",
      allowInput: true,
      maxDate: "today",
    });

    // Inputmask
    Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(input);
  });

  document
    .querySelectorAll("input[data-format='mm-yyyy']")
    .forEach(function (input) {
      flatpickr(input, {
        plugins: [
          new monthSelectPlugin({
            shorthand: false,
            dateFormat: "m-Y", // <-- Cómo se guarda
            altFormat: "F Y", // <-- Cómo se muestra (opcional)
            theme: "light",
            allowInput: true,
          }),
        ],
        allowInput: true,
      });

      // Inputmask para mm-yyyy
      Inputmask("99-9999", { placeholder: "mm-yyyy" }).mask(input);
    });

  camposAcentos.forEach(reemplazarAcentos);

  const selectCargo = document.getElementById("cargoActual");
  const otroCargoContainer = document.getElementById("otroCargoContainer");
  const otroCargoInput = document.getElementById("otroCargoInput");

  if (selectCargo) {
    selectCargo.addEventListener("change", function () {
      if (selectCargo.value === "Otro") {
        otroCargoContainer.classList.remove("hidden");
        otroCargoInput.setAttribute("required", true);
      } else {
        otroCargoContainer.classList.add("hidden");
        otroCargoInput.value = ""; // limpiar campo si cambia
        otroCargoInput.removeAttribute("required");
      }
    });
  }

  const selectCom = document.getElementById("comunidadIndigena");
  const otraComunidadContainer = document.getElementById(
    "otraComunidadContainer",
  );
  const otraComunidadInput = otraComunidadContainer.querySelector(
    'input[name="otraComunidad"]',
  );

  if (selectCom) {
    selectCom.addEventListener("change", function () {
      if (selectCom.value === "OTRO") {
        otraComunidadContainer.classList.remove("hidden");
        otraComunidadInput.setAttribute("required", true);
      } else {
        otraComunidadContainer.classList.add("hidden");
        otraComunidadInput.value = ""; // limpiar campo si cambia
        otraComunidadInput.removeAttribute("required");
      }
    });
  }

  const selectFam = document.getElementById("familiaLinguistica");
  const otraFamiliaContainer = document.getElementById("otraLenguaContainer");
  const otraFamiliaInput = otraFamiliaContainer.querySelector(
    'input[name="otraLengua"]',
  );

  if (selectFam) {
    selectFam.addEventListener("change", function () {
      if (selectFam.value === "OTRA") {
        otraFamiliaContainer.classList.remove("hidden");
        otraFamiliaInput.setAttribute("required", true);
      } else {
        otraFamiliaContainer.classList.add("hidden");
        otraFamiliaInput.value = ""; // limpiar campo si cambia
        otraFamiliaInput.removeAttribute("required");
      }
    });
  }

  const niveles = [
    { key: "Licenciatura", max: 3 },
    { key: "Maestria", max: 3 },
    { key: "Doctorado", max: 2 },
    { key: "Posdoctorado", max: 2 },
    { key: "Especialidad", max: 3 },
  ];
  const placeholders = {
    licenciatura: {
      titulo: "Ej. DERECHO, PSICOLOGIA, ADMINISTRACIÓN",
      institucion: "Ej. UNIVERSIDAD AUTONOMA DEL ESTADO DE HIDALGO",
      fecha: "dd-mm-yyyy",
      documento: "Diploma / Certificado",
      estatus: "Concluido / En curso / Titulado",
      cedula: "Ej. 1234567",
    },
    maestria: {
      titulo: "Ej. DERECHO CONSTITUCIONAL",
      fecha: "dd-mm-yyyy",
      institucion: "Ej. UNIVERSIDAD AUTONOMA DEL ESTADO DE HIDALGO",
      // ...y así para los demás si quieres
    },
    // similar para especialidad, doctorado...
  };
  niveles.forEach(({ key, max }) => {
    const tieneRadio = document.querySelectorAll(`input[name="tiene${key}"]`);
    const qtyContainer = document.getElementById(
      `${key.toLowerCase()}CantidadContainer`,
    );
    const qtyInput = document.getElementById(`${key.toLowerCase()}Cantidad`);
    const listaDiv = document.getElementById(`${key.toLowerCase()}Lista`);

    // Mostrar/ocultar el contador
    tieneRadio.forEach((radio) =>
      radio.addEventListener("change", () => {
        if (radio.value === "1") {
          qtyContainer.classList.remove("hidden");
          qtyInput.setAttribute("required", true);
        } else {
          qtyContainer.classList.add("hidden");
          qtyInput.removeAttribute("required");
          qtyInput.value = "";
          listaDiv.innerHTML = "";
        }
      }),
    );

    // Generar bloques según la cantidad
    qtyInput.addEventListener("input", () => {
      const n = Math.min(max, Math.max(0, parseInt(qtyInput.value || 0)));
      listaDiv.innerHTML = "";

      // Aquí sí usamos `key`, no `nivel`
      const ph = placeholders[key.toLowerCase()] || {};

      for (let i = 1; i <= n; i++) {
        listaDiv.insertAdjacentHTML(
          "beforeend",
          `
        <div class="border p-3 mt-4 mb-3">
          <p class="font-semibold mb-2">${key} ${i}</p>

          <label class="block mb-1">Nombre</label>

          <label class="block mb-1 text-red-500">Si te encuentras cursando colocar el semestre, cuatrimestre o bimestre</label>
          <input type="text"
                 name="${key.toLowerCase()}_titulo${i}"
                 placeholder="${ph.titulo || ""}"
                 required
                 class="border p-2 w-full rounded mb-2 acentos">

          <label class="block mb-1">Institución que expidió el documento</label>
          <input type="text"
                 name="${key.toLowerCase()}_institucion${i}"
                 placeholder="${ph.institucion || ""}"
                 required
                 class="border p-2 w-full rounded mb-2 acentos">
          
          <label class="block mb-1">Estatus</label>
          <select name="${key.toLowerCase()}_estatus${i}"
                  required
                  class="border p-2 w-full rounded mb-2">
            <option value="">Selecciona estatus</option>
            <option value="INCONCLUSO">INCONCLUSO</option>
            <option value="CURSANDO">CURSANDO</option>
            <option value="CONCLUIDO">CONCLUIDO</option>
            <option value="TITULADO">TITULADO</option>
          </select>

          <label class="block mb-1">Documento adquirido</label>
          <select name="${key.toLowerCase()}_documento${i}"
                  required
                  class="border p-2 w-full rounded mb-2">
            <option value=""> Selecciona documento</option>
            <option value="DIPLOMA">DIPLOMA</option>
            <option value="CONSTANCIA">CONSTANCIA</option>
            <option value="CERTIFICADO">CERTIFICADO</option>
            <option value="TITULO">TÍTULO</option>
            <option value="CEDULA">CÉDULA</option>
            <option value="TITULO Y CEDULA">TÍTULO Y CÉDULA</option>
          </select>

          <label class="block mb-1">Fecha de obtención del documento</label>
          <input type="text"
                 name="${key.toLowerCase()}_fecha${i}"
                 required
                placeholder="${ph.fecha || ""}"
                 data-format="dd-mm-yyyy"
                 class="border p-2 w-full rounded mb-2">
 
          <label class="block mb-1">Cédula profesional (Si no cuenta con una deje el campo vacio)</label>
          <input type="number"
                 name="${key.toLowerCase()}_cedula${i}"
                 placeholder="${ph.cedula || ""}"
                 class="border p-2 w-full rounded mb-2">
        </div>
      `,
        );
        configurarFechaInput();
        procesarCamposAcentos();
      }
    });
  });
});

function configurarFechaInput() {
  // Configuración para inputs con formato "d-m-Y"
  document
    .querySelectorAll("input[data-format='dd-mm-yyyy']")
    .forEach(function (input) {
      flatpickr(input, {
        dateFormat: "d-m-Y",
        allowInput: true,
        maxDate: "today",
      });

      Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(input);
    });

  // Configuración para inputs con formato "mm-yyyy"
  document
    .querySelectorAll("input[data-format='mm-yyyy']")
    .forEach(function (input) {
      flatpickr(input, {
        plugins: [
          new monthSelectPlugin({
            shorthand: false,
            dateFormat: "m-Y", // Cómo se guarda
            altFormat: "F Y", // Cómo se muestra (opcional)
            theme: "light",
            allowInput: true,
          }),
        ],
        allowInput: true,
      });

      Inputmask("99-9999", { placeholder: "mm-yyyy" }).mask(input);
    });
}

function procesarCamposAcentos() {
  const camposAcentos = document.querySelectorAll(".acentos");
  camposAcentos.forEach(reemplazarAcentos);
}

function reemplazarAcentos(campo) {
  campo.addEventListener("input", function () {
    //eliminar  losacentos
    //.value = this.value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    let valor = campo.value
      .replace(/ñ/gi, "ñ") // Mantener ñ existentes
      .normalize("NFD")
      // Convertir n + tilde a ñ (después de normalizar)
      .replace(/\u006E\u0303/gi, "ñ")
      // Eliminar otros diacríticos
      .replace(/[\u0300-\u036f]/g, "")
      // Permitir ñ/Ñ en el filtro
      .replace(/[^a-zA-Z0-9ñÑ\s\.]/g, "")
      .replace(/^\s+/, "")
      .replace(/\s{2,}/g, " ");

    campo.value = valor.toUpperCase();
  });
}

let experienciaIndex = 0;

function agregarExperiencia() {
  const container = document.getElementById("espacioExperiencia");

  const div = document.createElement("div");
  div.classList.add("mb-3", "p-4", "rounded", "border", "border-gray-300");

  div.innerHTML = `
<p class="font-semibold text-gray-800 mb-1">Experiencia ${experienciaIndex + 1}</p>
<label class="block text-gray-700 font-semibold">Periodo: Mes/Año de inicio</label>
<input type="text" name="experiencia_inicio${experienciaIndex}" data-format="mm-yyyy" required class="border w-full p-2 rounded-md bg-gray-50 mb-2" placeholder="MM-AAAA">

<label class="block text-gray-700 font-semibold"> Periodo: Mes/Año de fin</label>
<input type="month" name="experiencia_fin${experienciaIndex}" id="experiencia_fin${experienciaIndex}" data-format="mm-yyyy" required class="border w-full p-2 rounded-md bg-gray-50"  placeholder="MM-AAAA">


<input type="checkbox" name="experiencia_actual${experienciaIndex}" id="experiencia_actual${experienciaIndex}"> <span class="text-red-500"> A LA FECHA </span></input>

<label class="block text-gray-700 font-semibold">Institución</label>
  <select name="experiencia_institucion${experienciaIndex}"
          class="border w-full p-2 rounded-md bg-gray-50 mb-2">
    <option value="PJ">Poder Judicial del Estado de Hidalgo</option>
    <option value="OTRA">Otra</option>
  </select>

  <div id="adscripcion-container-${experienciaIndex}" class="mb-2">
    <label class="block text-gray-700 font-semibold">Adscripción</label>
    <input type="text" name="experiencia_adscripcion${experienciaIndex}"
           placeholder="EJ. PRIMERA SALA CIVIL Y FAMILIAR"
           class="acentos border w-full p-2 rounded-md bg-gray-50">
  </div>

  <div id="otra-inst-container-${experienciaIndex}" class="hidden mb-2">
    <label class="block text-gray-700 font-semibold">Nombre de la institución</label>
    <input type="text" name="experiencia_otrainst${experienciaIndex}"
           placeholder="EJ. BANCO AZTECA, S.A."
           class="acentos border w-full p-2 rounded-md bg-gray-50">
  </div>

<label class="block text-gray-700 font-semibold">Cargo ó Puesto desempeñado</label>
<input type="text" name="experiencia_puesto${experienciaIndex}" placeholder="Ej. SECRETARIA DE ESTUDIO Y CUENTA" required class="acentos border w-full p-2 rounded-md bg-gray-50 mb-2">

<label class="block text-gray-700 font-semibold">Campo de experiencia</label>
<input type="text" name="experiencia_campo${experienciaIndex}" placeholder="Ej. ELABORACION DE PROYECTOS DE SENTENCIA DE SEGUNDA INSTANCIA" required class="acentos border w-full p-2 rounded-md bg-gray-50 mb-2">


`;

  if (experienciaIndex < 100) {
    container.appendChild(div);

    // dentro de agregarExperiencia(), justo tras appendChild(div):
    const selectInst = div.querySelector(
      `select[name="experiencia_institucion${experienciaIndex}"]`,
    );
    const adscripContainer = div.querySelector(
      `#adscripcion-container-${experienciaIndex}`,
    );
    const otraContainer = div.querySelector(
      `#otra-inst-container-${experienciaIndex}`,
    );
    function toggleCampos(e) {
      if (e.target.value === "PJ") {
        adscripContainer.classList.remove("hidden");
        otraContainer.classList.add("hidden");
      } else {
        adscripContainer.classList.add("hidden");
        otraContainer.classList.remove("hidden");
      }
    }

    // Inicializa según el valor por defecto
    toggleCampos({ target: selectInst });

    // Escucha cambios
    selectInst.addEventListener("change", toggleCampos);

    const finInput = document.getElementById(
      `experiencia_fin${experienciaIndex}`,
    );
    const actualCb = document.getElementById(
      `experiencia_actual${experienciaIndex}`,
    );

    actualCb.addEventListener("change", () => {
      if (actualCb.checked) {
        // ocultar visualmente y quitar validación
        finInput.classList.add("hidden");
        finInput.removeAttribute("required");
        finInput.value = ""; // limpia cualquier valor previo
      } else {
        // volver a mostrar y requerir
        finInput.classList.remove("hidden");
        finInput.setAttribute("required", "true");
      }
    });

    experienciaIndex++;
    document.getElementById("experienciaTotal").value = experienciaIndex;
  }

  document
    .querySelectorAll("input[data-format='mm-yyyy']")
    .forEach(function (input) {
      flatpickr(input, {
        plugins: [
          new monthSelectPlugin({
            shorthand: false,
            dateFormat: "m-Y", // <-- Cómo se guarda
            altFormat: "F Y", // <-- Cómo se muestra (opcional)
            theme: "light",
            allowInput: true,
          }),
        ],
        allowInput: true,
      });

      // Inputmask para mm-yyyy
      Inputmask("99-9999", { placeholder: "mm-yyyy" }).mask(input);
    });

  const camposAcentos = document.querySelectorAll(".acentos");
  camposAcentos.forEach(reemplazarAcentos);
}

function quitarExperiencia() {
  const container = document.getElementById("espacioExperiencia");
  container.removeChild(container.lastChild);
  experienciaIndex = Math.max(0, experienciaIndex - 1);
  document.getElementById("experienciaTotal").value = experienciaIndex;
}

let actualizacionIndex = 0;

function agregarActualizacion() {
  const container = document.getElementById("espacioActualizacion");

  if (actualizacionIndex >= 10) return;

  const div = document.createElement("div");
  div.classList.add(
    "mb-4",
    "p-4",
    "rounded",
    "border",
    "border-gray-300",
    "bg-gray-50",
  );

  div.innerHTML = `
    <p class="font-semibold text-gray-800 mb-2">Actualización ${actualizacionIndex + 1}</p>

    <label class="block text-gray-700 font-semibold">Tema</label>
    <input type="text" name="actualizacion_tema${actualizacionIndex}" class="acentos border w-full p-2 rounded-md bg-white mb-2" required>

    <label class="block text-gray-700 font-semibold">Fecha (día/mes/año)</label>
    <input type="text" name="actualizacion_fecha${actualizacionIndex}" data-format="dd-mm-yyyy" class="border w-full p-2 rounded-md bg-white mb-2" placeholder="DD-MM-AAAA" required>

    <label class="block text-gray-700 font-semibold">Institución</label>
    <input type="text" name="actualizacion_institucion${actualizacionIndex}" class="acentos border w-full p-2 rounded-md bg-white mb-2" required>

    <label class="block text-gray-700 font-semibold">Documento recibido</label>
     <select  class="acentos border w-full p-2 rounded-md bg-white mb-2" name="actualizacion_documento${actualizacionIndex}" required>
      <option value="">Seleccione documento</option>
      <option value="CERTIFICADO">CERTIFICADO</option>
      <option value="DIPLOMA">DIPLOMA</option>
      <option value="CONSTANCIA">CONSTANCIA</option>
      <option value="RECONOCIMIENTO">RECONOCIMIENTO</option>
    </select>
  
    `;

  container.appendChild(div);

  // Activar flatpickr + inputmask en el nuevo campo de fecha
  const inputFecha = div.querySelector(
    `[name="actualizacion_fecha${actualizacionIndex}"]`,
  );
  flatpickr(inputFecha, {
    dateFormat: "d-m-Y", // Día-Mes-Año
    allowInput: true,
    maxDate: "today",
  });

  Inputmask("99-99-9999", { placeholder: "dd-mm-yyyy" }).mask(inputFecha);

  // Activar el filtro de acentos
  const camposAcentos = div.querySelectorAll(".acentos");
  camposAcentos.forEach(reemplazarAcentos);

  actualizacionIndex++;
  document.getElementById("actualizacionTotal").value = actualizacionIndex;
}

function quitarActualizacion() {
  const container = document.getElementById("espacioActualizacion");
  if (container.lastChild) {
    container.removeChild(container.lastChild);
    actualizacionIndex = Math.max(0, actualizacionIndex - 1);
    document.getElementById("actualizacionTotal").value = actualizacionIndex;
  }
}

function validarEscolaridad(prefix, cantidadMaxima, nombreVisible) {
  for (let i = 1; i <= cantidadMaxima; i++) {
    const titulo = document.querySelector(`[name="${prefix}_titulo${i}"]`);
    const fecha = document.querySelector(`[name="${prefix}_fecha${i}"]`);
    const inst = document.querySelector(`[name="${prefix}_institucion${i}"]`);
    const documento = document.querySelector(
      `[name="${prefix}_documento${i}"]`,
    );
    const estatus = document.querySelector(`[name="${prefix}_estatus${i}"]`);
    const cedula = document.querySelector(`[name="${prefix}_cedula${i}"]`);

    // Si alguno de los campos obligatorios no existe, saltamos
    if (!titulo || !fecha || !inst || !documento || !estatus) continue;

    const tieneAlgunDato =
      titulo.value.trim() ||
      fecha.value.trim() ||
      inst.value.trim() ||
      documento.value ||
      estatus.value ||
      (cedula && cedula.value.trim());

    const incompleto =
      !titulo.value.trim() ||
      !fecha.value.trim() ||
      !inst.value.trim() ||
      !documento.value ||
      !estatus.value;

    if (tieneAlgunDato && incompleto) {
      alert(
        `⚠️ Completa todos los campos obligatorios de ${nombreVisible} ${i} o deja todos vacíos.`,
      );
      if (!titulo.value.trim()) titulo.focus();
      else if (!fecha.value.trim()) fecha.focus();
      else if (!inst.value.trim()) inst.focus();
      else if (!documento.value) documento.focus();
      else if (!estatus.value) estatus.focus();
      return false;
    }
  }
  return true;
}

function validarActualizaciones(cantidadMaxima) {
  for (let i = 0; i < cantidadMaxima; i++) {
    const tema = document.querySelector(`[name="actualizacion_tema${i}"]`);
    const fecha = document.querySelector(`[name="actualizacion_fecha${i}"]`);
    const inst = document.querySelector(
      `[name="actualizacion_institucion${i}"]`,
    );
    const doc = document.querySelector(`[name="actualizacion_documento${i}"]`);

    if (!tema || !fecha || !inst || !doc) continue;

    const hayAlgo =
      tema.value.trim() ||
      fecha.value.trim() ||
      inst.value.trim() ||
      doc.value.trim();
    const incompleto =
      !tema.value.trim() ||
      !fecha.value.trim() ||
      !inst.value.trim() ||
      !doc.value.trim();

    if (hayAlgo && incompleto) {
      alert(
        `⚠️ Completa todos los campos de la actualización profesional ${i + 1} o deja todos vacíos.`,
      );
      if (!tema.value.trim()) tema.focus();
      else if (!fecha.value.trim()) fecha.focus();
      else if (!inst.value.trim()) inst.focus();
      else doc.focus();
      return false;
    }
  }
  return true;
}
