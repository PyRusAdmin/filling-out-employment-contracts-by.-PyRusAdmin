async function handleAction(actionId) {
    try {
        const response = await fetch('/action', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
            },
            body: `user_input=${actionId}`,
        });

        if (response.ok) {
            // Если сервер вернул Redirect (код 303), перенаправляем на URL из ответа
            const url = response.url;
            window.location.href = url;
        } else {
            alert('Ошибка при выполнении действия!');
        }
    } catch (error) {
        console.error('Ошибка:', error);
        alert('Произошла ошибка!');
    }
}