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
    const filterButtons = document.querySelectorAll('.filter-btn');
    
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
            const priorityOrder = {urgent: 0, alta: 1, normal: 2};
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
                urgent: 'Urgente'
            }[task.priority];
            
            taskItem.innerHTML = `
                <input type="checkbox" class="task-checkbox" ${task.completed ? 'checked' : ''} data-id="${task.id}">
                <span class="task-text">${escapeHTML(task.text)}</span>
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
            {id: 2, text: 'Crear aplicación web', priority: 'urgent', completed: false, createdAt: new Date().toISOString()},
            {id: 3, text: 'Diseñar interfaz responsive', priority: 'normal', completed: false, createdAt: new Date().toISOString()},
            {id: 4, text: 'Probar la aplicación', priority: 'alta', completed: false, createdAt: new Date().toISOString()}
        ];
        saveTasks();
        renderTasks();
    }
});