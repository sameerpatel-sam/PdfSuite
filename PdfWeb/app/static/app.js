const cards = document.querySelectorAll('.tool-card');
const forms = document.querySelectorAll('.tool-form');
const loading = document.getElementById('loading');
const toast = document.getElementById('toast');

function showToast(msg) {
  toast.textContent = msg;
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), 3500);
}

cards.forEach(card => {
  card.addEventListener('click', () => {
    const tool = card.getAttribute('data-tool');
    forms.forEach(f => f.classList.add('hidden'));
    const target = document.getElementById(`form-${tool}`);
    if (target) target.classList.remove('hidden');
    window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
  });
});

forms.forEach(form => {
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const endpoint = form.getAttribute('data-endpoint');
    const fd = new FormData(form);

    // Client-side file validations
    for (const [key, value] of fd.entries()) {
      if (value instanceof File && value.size === 0) {
        showToast('Please select a file');
        return;
      }
    }

    loading.classList.remove('hidden');
    try {
      const res = await fetch(endpoint, { method: 'POST', body: fd });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'Request failed');
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const disposition = res.headers.get('Content-Disposition') || '';
      const match = /filename=([^;]+)$/i.exec(disposition);
      a.href = url;
      a.download = match ? match[1] : 'download';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      showToast(e.message);
    } finally {
      loading.classList.add('hidden');
    }
  });
});
