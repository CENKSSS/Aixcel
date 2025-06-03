function openModal(contentHTML) {
    let modal = document.getElementById('customModal');
    if (!modal) {
        modal = document.createElement('div');
        modal.id = 'customModal';
        modal.style.position = 'fixed';
        modal.style.top = '0'; modal.style.left = '0';
        modal.style.width = '100vw'; modal.style.height = '100vh';
        modal.style.background = 'rgba(0,0,0,0.4)';
        modal.style.display = 'flex'; modal.style.justifyContent = 'center'; modal.style.alignItems = 'center';
        modal.style.zIndex = '10000';
        modal.innerHTML = `
            <div style="background:#fff;padding:24px;min-width:300px;min-height:200px;position:relative;border-radius:10px;box-shadow:0 4px 32px #0002;">
                <button onclick="closeModal()" style="position:absolute;top:8px;right:12px;font-size:24px;background:none;border:none;cursor:pointer;">&times;</button>
                <div id="modalInnerContent">${contentHTML || ""}</div>
            </div>`;
        document.body.appendChild(modal);
    } else {
        document.getElementById('modalInnerContent').innerHTML = contentHTML || "";
        modal.style.display = 'flex';
    }
}
function closeModal() {
    const modal = document.getElementById('customModal');
    if (modal) modal.style.display = 'none';
}
window.openModal = openModal;
window.closeModal = closeModal;
