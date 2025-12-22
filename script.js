document.addEventListener('DOMContentLoaded', function() {
    // Elementos del DOM
    const taskInput = document.getElementById('taskInput');
    const prioritySelect = document.getElementById('prioritySelect');
    const addTaskBtn = document.getElementById('addTaskBtn');
    const taskList = document.getElementById('taskList');
    const totalTasks = document.getElementById('totalTasks');
    const completedTasks = document.getElementById('completedTasks');
    const pendingTasks = document.getElementById('pendingTasks');
    const clearCompletedBtn = document.getElementById('clearCompletedBtn');
    const saveTasksBtn = document.getElementById('saveTasksBtn');
    const importExcelBtn = document.getElementById('importExcelBtn');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const importTasksBtn = document.getElementById('importTasksBtn');
    const exportTasksBtn = document.getElementById('exportTasksBtn');
    const filterButtons = document.querySelectorAll('.filter-btn');
    const fileInput = document.getElementById('fileInput');
    
    // Variables para importación
    let importMode = 'replace'; // 'replace', 'merge', 'add'
    
    // Cargar tareas desde localStorage
    let tasks = JSON.parse(localStorage.getItem('tasks')) || [];
    let currentFilter = 'all';
    
    // Inicializar la aplicación
    updateStats();
    renderTasks();
    
    // Agregar nueva tarea
    addTaskBtn.addEventListener('click', addTask);
    taskInput.addEventListener('keypress', function(e) {
        if (e.key === 'Enter') addTask();
    });
    
    // Filtrar tareas
    filterButtons.forEach(button => {
        button.addEventListener('click', function() {
            // Remover clase active de todos los botones
            filterButtons.forEach(btn => btn.classList.remove('active'));
            // Agregar clase active al botón clickeado
            this.classList.add('active');
            currentFilter = this.getAttribute('data-filter');
            renderTasks();
        });
    });
    
    // Limpiar tareas completadas
    clearCompletedBtn.addEventListener('click', function() {
        tasks = tasks.filter(task => !task.completed);
        saveTasks();
        renderTasks();
        showNotification('Tareas completadas eliminadas', 'info');
    });
    
    // Guardar tareas manualmente
    saveTasksBtn.addEventListener('click', function() {
        saveTasks();
        showNotification('Tareas guardadas exitosamente', 'success');
    });
    
    // Exportar a Excel
    exportExcelBtn.addEventListener('click', exportToExcel);
    exportTasksBtn.addEventListener('click', exportToExcel);
    
    // Importar desde Excel
    importExcelBtn.addEventListener('click', showImportModal);
    importTasksBtn.addEventListener('click', showImportModal);
    
    // Manejar selección de archivo
    fileInput.addEventListener('change', handleFileSelect);
    
    // Función para mostrar modal de importación
    function showImportModal() {
        // Crear modal si no existe
        let modal = document.getElementById('importModal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'importModal';
            modal.className = 'modal';
            modal.innerHTML = `
                <div class="modal-content">
                    <h2><i class="fas fa-file-import"></i> Importar desde Excel</h2>
                    <p>Selecciona cómo quieres importar las tareas:</p>
                    
                    <div class="modal-options">
                        <div class="modal-option" data-mode="replace">
                            <i class="fas fa-sync-alt"></i>
                            <div>
                                <h3>Reemplazar todo</h3>
                                <p>Elimina todas las tareas actuales y carga las del archivo Excel</p>
                            </div>
                        </div>
                        
                        <div class="modal-option" data-mode="merge">
                            <i class="fas fa-blender"></i>
                            <div>
                                <h3>Combinar</h3>
                                <p>Mantiene las tareas actuales y añade las del archivo (sin duplicados)</p>
                            </div>
                        </div>
                        
                        <div class="modal-option" data-mode="add">
                            <i class="fas fa-plus-circle"></i>
                            <div>
                                <h3>Añadir nuevas</h3>
                                <p>Añade todas las tareas del archivo Excel a las existentes</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="modal-buttons">
                        <button class="modal-btn modal-btn-cancel" id="cancelImportBtn">Cancelar</button>
                        <button class="modal-btn modal-btn-confirm" id="confirmImportBtn" disabled>
                            <i class="fas fa-file-import"></i> Seleccionar Archivo
                        </button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);
            
            // Agregar event listeners a las opciones
            const options = modal.querySelectorAll('.modal-option');
            options.forEach(option => {
                option.addEventListener('click', function() {
                    // Remover selección de todas las opciones
                    options.forEach(opt => opt.style.borderColor = '#e0e0e0');
                    // Seleccionar esta opción
                    this.style.borderColor = '#2575fc';
                    importMode = this.getAttribute('data-mode');
                    modal.querySelector('#confirmImportBtn').disabled = false;
                });
            });
            
            // Botón cancelar
            modal.querySelector('#cancelImportBtn').addEventListener('click', function() {
                modal.style.display = 'none';
            });
            
            // Botón confirmar
            modal.querySelector('#confirmImportBtn').addEventListener('click', function() {
                modal.style.display = 'none';
                fileInput.click();
            });
            
            // Cerrar modal al hacer clic fuera
            modal.addEventListener('click', function(e) {
                if (e.target === modal) {
                    modal.style.display = 'none';
                }
            });
        }
        
        // Mostrar modal
        modal.style.display = 'flex';
    }
    
    // Función para manejar selección de archivo
    function handleFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        // Verificar extensión
        const fileName = file.name.toLowerCase();
        if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
            showNotification('Por favor, selecciona un archivo Excel (.xlsx o .xls)', 'error');
            fileInput.value = '';
            return;
        }
        
        const reader = new FileReader();
        
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Obtener la primera hoja
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const excelData = XLSX.utils.sheet_to_json(firstSheet);
                
                // Procesar datos del Excel
                processImportedData(excelData, file.name);
                
                // Limpiar input file
                fileInput.value = '';
            } catch (error) {
                console.error('Error al leer el archivo Excel:', error);
                showNotification('Error al leer el archivo Excel. Verifica el formato.', 'error');
                fileInput.value = '';
            }
        };
        
        reader.onerror = function() {
            showNotification('Error al leer el archivo', 'error');
            fileInput.value = '';
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // Función para procesar datos importados
    function processImportedData(excelData, fileName) {
        if (!excelData || excelData.length === 0) {
            showNotification('El archivo Excel no contiene datos válidos', 'error');
            return;
        }
        
        const importedTasks = [];
        let skippedTasks = 0;
        
        // Procesar cada fila del Excel
        excelData.forEach((row, index) => {
            // Validar datos mínimos
            if (!row.Tarea || row.Tarea.trim() === '') {
                skippedTasks++;
                return;
            }
            
            // Convertir prioridad de texto a valor interno
            let priority = 'normal';
            if (row.Prioridad) {
                const prioridad = row.Prioridad.toLowerCase();
                if (prioridad.includes('alta') || prioridad.includes('high')) priority = 'alta';
                else if (prioridad.includes('urgente') || prioridad.includes('urgent')) priority = 'urgente';
            }
            
            // Convertir estado de texto a booleano
            let completed = false;
            if (row.Estado) {
                const estado = row.Estado.toLowerCase();
                if (estado.includes('completada') || estado.includes('completed') || 
                    estado.includes('terminada') || estado.includes('done')) {
                    completed = true;
                }
            }
            
            // Generar nuevo ID para evitar conflictos
            const newId = Date.now() + index;
            
            // Intentar parsear fecha o usar fecha actual
            let createdAt = new Date().toISOString();
            if (row['Fecha de Creación']) {
                try {
                    const parsedDate = new Date(row['Fecha de Creación']);
                    if (!isNaN(parsedDate.getTime())) {
                        createdAt = parsedDate.toISOString();
                    }
                } catch (e) {
                    // Usar fecha actual si hay error
                }
            }
            
            // Crear objeto de tarea
            const task = {
                id: newId,
                text: row.Tarea.trim(),
                priority: priority,
                completed: completed,
                createdAt: createdAt,
                imported: true,
                originalId: row.ID || null
            };
            
            importedTasks.push(task);
        });
        
        // Aplicar el modo de importación seleccionado
        let message = '';
        let taskCount = 0;
        
        switch(importMode) {
            case 'replace':
                tasks = importedTasks;
                message = `Todas las tareas reemplazadas. ${importedTasks.length} tareas importadas.`;
                taskCount = importedTasks.length;
                break;
                
            case 'merge':
                // Filtrar tareas duplicadas (mismo texto)
                const existingTexts = new Set(tasks.map(t => t.text.toLowerCase()));
                const newTasks = importedTasks.filter(t => !existingTexts.has(t.text.toLowerCase()));
                
                tasks = [...tasks, ...newTasks];
                message = `${newTasks.length} nuevas tareas añadidas (${skippedTasks + importedTasks.length - newTasks.length} duplicadas omitidas).`;
                taskCount = newTasks.length;
                break;
                
            case 'add':
                tasks = [...tasks, ...importedTasks];
                message = `${importedTasks.length} tareas añadidas.`;
                taskCount = importedTasks.length;
                break;
        }
        
        // Guardar y renderizar
        saveTasks();
        renderTasks();
        
        // Mostrar resumen de importación
        showImportSummary(message, taskCount, skippedTasks, fileName);
    }
    
    // Función para mostrar resumen de importación
    function showImportSummary(message, importedCount, skippedCount, fileName) {
        const summaryModal = document.createElement('div');
        summaryModal.className = 'modal';
        summaryModal.style.display = 'flex';
        summaryModal.innerHTML = `
            <div class="modal-content">
                <h2><i class="fas fa-check-circle" style="color: #28a745;"></i> Importación Completada</h2>
                
                <div style="margin: 20px 0;">
                    <p><strong>Archivo:</strong> ${fileName}</p>
                    <p><strong>Modo:</strong> ${getImportModeName(importMode)}</p>
                    <p><strong>Tareas importadas:</strong> ${importedCount}</p>
                    ${skippedCount > 0 ? `<p><strong>Filas omitidas:</strong> ${skippedCount}</p>` : ''}
                    <p><strong>Total de tareas ahora:</strong> ${tasks.length}</p>
                </div>
                
                <div style="background: #f8f9fa; padding: 15px; border-radius: 10px; margin: 20px 0;">
                    <p style="margin: 0; color: #333;">${message}</p>
                </div>
                
                <div class="modal-buttons">
                    <button class="modal-btn modal-btn-confirm" id="closeSummaryBtn">
                        <i class="fas fa-check"></i> Aceptar
                    </button>
                </div>
            </div>
        `;
        
        document.body.appendChild(summaryModal);
        
        // Cerrar modal
        summaryModal.querySelector('#closeSummaryBtn').addEventListener('click', function() {
            summaryModal.remove();
        });
        
        summaryModal.addEventListener('click', function(e) {
            if (e.target === summaryModal) {
                summaryModal.remove();
            }
        });
    }
    
    // Función para obtener nombre del modo de importación
    function getImportModeName(mode) {
        const names = {
            'replace': 'Reemplazar todo',
            'merge': 'Combinar',
            'add': 'Añadir nuevas'
        };
        return names[mode] || mode;
    }
    
    // Función para agregar tarea
    function addTask() {
        const text = taskInput.value.trim();
        const priority = prioritySelect.value;
        
        if (text === '') {
            showNotification('Por favor, escribe una tarea', 'error');
            taskInput.focus();
            return;
        }
        
        const newTask = {
            id: Date.now(),
            text: text,
            priority: priority,
            completed: false,
            createdAt: new Date().toISOString()
        };
        
        tasks.push(newTask);
        taskInput.value = '';
        taskInput.focus();
        
        saveTasks();
        renderTasks();
        showNotification('Tarea agregada exitosamente', 'success');
    }
    
    // Función para renderizar tareas
    function renderTasks() {
        // Filtrar tareas según el filtro actual
        let filteredTasks = tasks;
        if (currentFilter === 'pending') {
            filteredTasks = tasks.filter(task => !task.completed);
        } else if (currentFilter === 'completed') {
            filteredTasks = tasks.filter(task => task.completed);
        }
        
        // Limpiar lista
        taskList.innerHTML = '';
        
        // Mostrar mensaje si no hay tareas
        if (filteredTasks.length === 0) {
            let message = '';
            if (currentFilter === 'all') message = 'No hay tareas. ¡Agrega una nueva!';
            else if (currentFilter === 'pending') message = 'No hay tareas pendientes';
            else if (currentFilter === 'completed') message = 'No hay tareas completadas';
            
            taskList.innerHTML = `
                <div class="empty-state">
                    <i class="fas fa-clipboard-list" style="font-size: 3rem; color: #ccc; margin-bottom: 10px;"></i>
                    <p style="color: #999; text-align: center;">${message}</p>
                </div>
            `;
            return;
        }
        
        // Ordenar tareas: urgentes primero, luego completadas al final
        filteredTasks.sort((a, b) => {
            // Ordenar por prioridad
            const priorityOrder = {urgente: 0, alta: 1, normal: 2};
            if (priorityOrder[a.priority] !== priorityOrder[b.priority]) {
                return priorityOrder[a.priority] - priorityOrder[b.priority];
            }
            // Luego por estado (pendientes primero)
            if (a.completed !== b.completed) {
                return a.completed ? 1 : -1;
            }
            // Finalmente por fecha de creación
            return new Date(b.createdAt) - new Date(a.createdAt);
        });
        
        // Renderizar cada tarea
        filteredTasks.forEach(task => {
            const taskItem = document.createElement('li');
            taskItem.className = `task-item ${task.priority}-priority ${task.completed ? 'completed' : ''}`;
            
            // Obtener texto de prioridad en español
            const priorityText = {
                normal: 'Normal',
                alta: 'Alta',
                urgente: 'Urgente'
            }[task.priority];
            
            // Agregar indicador de tarea importada
            const importIndicator = task.imported ? 
                '<span class="imported-indicator" title="Importada desde Excel"><i class="fas fa-file-import"></i></span>' : '';
            
            taskItem.innerHTML = `
                <input type="checkbox" class="task-checkbox" ${task.completed ? 'checked' : ''} data-id="${task.id}">
                <span class="task-text">${escapeHTML(task.text)}</span>
                ${importIndicator}
                <span class="task-priority priority-${task.priority}">${priorityText}</span>
                <button class="delete-task" data-id="${task.id}">
                    <i class="fas fa-trash"></i>
                </button>
            `;
            
            taskList.appendChild(taskItem);
        });
        
        // Agregar event listeners a los checkboxes y botones de eliminar
        document.querySelectorAll('.task-checkbox').forEach(checkbox => {
            checkbox.addEventListener('change', toggleTask);
        });
        
        document.querySelectorAll('.delete-task').forEach(button => {
            button.addEventListener('click', deleteTask);
        });
        
        updateStats();
    }
    
    // Función para alternar estado de tarea
    function toggleTask(e) {
        const taskId = parseInt(e.target.getAttribute('data-id'));
        const task = tasks.find(t => t.id === taskId);
        
        if (task) {
            task.completed = e.target.checked;
            saveTasks();
            renderTasks();
            
            const message = task.completed ? 'Tarea completada' : 'Tarea marcada como pendiente';
            showNotification(message, 'info');
        }
    }
    
    // Función para eliminar tarea
    function deleteTask(e) {
        const taskId = parseInt(e.currentTarget.getAttribute('data-id'));
        
        if (confirm('¿Estás seguro de que quieres eliminar esta tarea?')) {
            tasks = tasks.filter(task => task.id !== taskId);
            saveTasks();
            renderTasks();
            showNotification('Tarea eliminada', 'error');
        }
    }
    
    // Función para actualizar estadísticas
    function updateStats() {
        const total = tasks.length;
        const completed = tasks.filter(task => task.completed).length;
        const pending = total - completed;
        
        totalTasks.textContent = total;
        completedTasks.textContent = completed;
        pendingTasks.textContent = pending;
    }
    
    // Función para guardar tareas en localStorage
    function saveTasks() {
        localStorage.setItem('tasks', JSON.stringify(tasks));
        updateStats();
    }
    
    // Función para exportar tareas a Excel
    function exportToExcel() {
        if (tasks.length === 0) {
            showNotification('No hay tareas para exportar', 'error');
            return;
        }
        
        // Preparar los datos para Excel
        const excelData = tasks.map(task => {
            // Formatear fecha
            const date = new Date(task.createdAt);
            const formattedDate = date.toLocaleDateString('es-ES', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
            
            // Convertir prioridad a texto completo
            const priorityText = {
                normal: 'Normal',
                alta: 'Alta',
                urgente: 'Urgente'
            }[task.priority];
            
            // Convertir estado a texto
            const statusText = task.completed ? 'Completada' : 'Pendiente';
            
            return {
                'ID': task.id,
                'Tarea': task.text,
                'Prioridad': priorityText,
                'Estado': statusText,
                'Fecha de Creación': formattedDate
            };
        });
        
        // Crear hoja de trabajo
        const worksheet = XLSX.utils.json_to_sheet(excelData);
        
        // Ajustar el ancho de las columnas
        const columnWidths = [
            { wch: 10 }, // ID
            { wch: 40 }, // Tarea
            { wch: 12 }, // Prioridad
            { wch: 12 }, // Estado
            { wch: 20 }  // Fecha de Creación
        ];
        worksheet['!cols'] = columnWidths;
        
        // Crear libro de trabajo
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Tareas');
        
        // Agregar hoja con estadísticas
        const statsData = [
            ['ESTADÍSTICAS', ''],
            ['Total de Tareas', tasks.length],
            ['Tareas Completadas', tasks.filter(t => t.completed).length],
            ['Tareas Pendientes', tasks.filter(t => !t.completed).length],
            ['', ''],
            ['Fecha de Exportación', new Date().toLocaleDateString('es-ES')]
        ];
        
        const statsWorksheet = XLSX.utils.aoa_to_sheet(statsData);
        XLSX.utils.book_append_sheet(workbook, statsWorksheet, 'Estadísticas');
        
        // Generar el archivo Excel
        const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        
        // Descargar el archivo
        const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        
        // Nombre del archivo con fecha actual
        const date = new Date();
        const dateStr = date.toISOString().split('T')[0];
        link.download = `lista_tareas_${dateStr}.xlsx`;
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        
        showNotification(`Se exportaron ${tasks.length} tareas a Excel`, 'success');
    }
    
    // Función para mostrar notificaciones
    function showNotification(message, type) {
        // Eliminar notificación anterior si existe
        const existingNotification = document.querySelector('.notification');
        if (existingNotification) {
            existingNotification.remove();
        }
        
        // Crear nueva notificación
        const notification = document.createElement('div');
        notification.className = `notification notification-${type}`;
        notification.innerHTML = `
            <i class="fas ${type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-exclamation-circle' : 'fa-info-circle'}"></i>
            <span>${message}</span>
        `;
        
        // Estilos para la notificación
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            background: ${type === 'success' ? '#4CAF50' : type === 'error' ? '#f44336' : '#2196F3'};
            color: white;
            border-radius: 5px;
            display: flex;
            align-items: center;
            gap: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            z-index: 1000;
            animation: slideIn 0.3s ease;
        `;
        
        document.body.appendChild(notification);
        
        // Auto-eliminar después de 3 segundos
        setTimeout(() => {
            notification.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => notification.remove(), 300);
        }, 3000);
        
        // Agregar estilos de animación si no existen
        if (!document.querySelector('#notification-styles')) {
            const style = document.createElement('style');
            style.id = 'notification-styles';
            style.textContent = `
                @keyframes slideIn {
                    from { transform: translateX(100%); opacity: 0; }
                    to { transform: translateX(0); opacity: 1; }
                }
                @keyframes slideOut {
                    from { transform: translateX(0); opacity: 1; }
                    to { transform: translateX(100%); opacity: 0; }
                }
            `;
            document.head.appendChild(style);
        }
    }
    
    // Función para escapar HTML (seguridad básica)
    function escapeHTML(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    // Cargar algunas tareas de ejemplo si no hay ninguna
    if (tasks.length === 0) {
        tasks = [
            {id: 1, text: 'Aprender JavaScript', priority: 'alta', completed: true, createdAt: new Date().toISOString()},
            {id: 2, text: 'Crear aplicación web', priority: 'urgente', completed: false, createdAt: new Date().toISOString()},
            {id: 3, text: 'Diseñar interfaz responsive', priority: 'normal', completed: false, createdAt: new Date().toISOString()},
            {id: 4, text: 'Probar la aplicación', priority: 'alta', completed: false, createdAt: new Date().toISOString()}
        ];
        saveTasks();
        renderTasks();
    }
});