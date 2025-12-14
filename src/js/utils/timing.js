export function debouncePromise(fn, delay = 300) {
    let timer = null;
    let pending = [];
    let lastArgs = null;

    return (...args) => {
        lastArgs = args;

        return new Promise((resolve, reject) => {
            pending.push({ resolve, reject });
            if (timer) clearTimeout(timer);

            timer = setTimeout(async () => {
                const queued = [...pending];
                pending = [];
                timer = null;

                try {
                    const result = await fn(...lastArgs);
                    queued.forEach(({ resolve: res }) => res(result));
                } catch (e) {
                    queued.forEach(({ reject: rej }) => rej(e));
                }
            }, delay);
        });
    };
}

export function debounce(fn, delay = 200) {
    let timer = null;
    return (...args) => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
}
