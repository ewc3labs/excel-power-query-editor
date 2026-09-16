/**
 * A file watcher that falls back to polling when native watching fails, and stops cleanly when
 * polling fails too.
 *
 * WHY THIS EXISTS. On a mapped network drive, chokidar's native watcher reported `ready`, errored 2ms
 * later with `UNKNOWN: unknown error, watch`, and never fired again - a watcher that says ready and is
 * deaf. See docs/project/slices/PQ-36_File_Watcher_On_Network_Drives.md.
 *
 * We do NOT classify the drive first. Node cannot tell whether `P:\` is a local disk or a mapped
 * share, and a UNC test catches only paths that were never abbreviated to a letter. So the failure
 * decides: native first, polling on the first native error, stopped on a polling error.
 *
 * WHY A SEPARATE MODULE. The first version of this lived inside extension.ts and was declared
 * untestable because "faking the error would test the fake". That was wrong, and a review said so:
 * what needs testing here is the STATE MACHINE - swap once, repoint the holder, ignore stale errors,
 * stop on the second failure - and an injected factory tests exactly that. Whether polling actually
 * works on somebody's file server is a different question, and only a real share answers it.
 */

/** The part of a watcher this module depends on. chokidar's FSWatcher satisfies it. */
export interface WatcherHandle {
	on(event: 'error', listener: (error: unknown) => void): unknown;
	close(): Promise<void>;
}

export type WatcherState = 'native' | 'polling' | 'stopped';

export interface ResilientWatcherOptions<W extends WatcherHandle> {
	/** Build a watcher. `true` means polling. */
	create(usePolling: boolean): W;
	/** Attach every listener except `error`, which this module owns. Called once per watcher built. */
	attach(watcher: W, usePolling: boolean): void;
	/** The live watcher was replaced. Repoint anything holding the old one, or it will close the wrong one. */
	onReplaced(watcher: W): void;
	/** Native and polling have both failed, or polling was all there was. The watcher is already closed. */
	onUnwatchable(error: unknown): void;
	log(message: string, level: 'info' | 'error'): void;
	/** Start already polling - dev containers. Then a failure has nothing to fall back to. */
	startPolling?: boolean;
}

export interface ResilientWatcher<W extends WatcherHandle> {
	/** The live watcher. Changes once, on fallback. */
	readonly current: W;
	readonly state: WatcherState;
	/** Stop watching. Errors that arrive afterwards are ignored. */
	close(): Promise<void>;
}

export function createResilientWatcher<W extends WatcherHandle>(opts: ResilientWatcherOptions<W>): ResilientWatcher<W> {
	let state: WatcherState = opts.startPolling ? 'polling' : 'native';

	const start = (usePolling: boolean): W => {
		const w = opts.create(usePolling);
		opts.attach(w, usePolling);
		w.on('error', (error) => { void onError(w, error); });
		return w;
	};

	let current = start(state === 'polling');

	async function onError(source: W, error: unknown): Promise<void> {
		// A replaced or closed watcher can still emit - chokidar reports errors asynchronously, and a
		// failing native watcher rarely fails exactly once. Only the live watcher decides anything, or
		// a late error from the dead one would start a second fallback or stop the working one.
		if (source !== current || state === 'stopped') { return; }

		opts.log(`Watcher error: ${error}`, 'error');

		if (state === 'polling') {
			state = 'stopped';
			await source.close().catch(() => { /* it has already failed; closing is courtesy */ });
			opts.onUnwatchable(error);
			return;
		}

		// Native failed. Claim the transition BEFORE any await, so a second error in the same tick finds
		// `source !== current` and does nothing, rather than building a second polling watcher.
		state = 'polling';
		opts.log('Native file watching failed - this usually means a network drive or UNC path. Retrying with polling.', 'info');
		const replacement = start(true);
		current = replacement;
		opts.onReplaced(replacement);
		// Replacement first, then close: there is no moment with nothing watching.
		await source.close().catch(() => { /* already dead */ });
	}

	return {
		get current() { return current; },
		get state() { return state; },
		async close() {
			state = 'stopped';
			await current.close();
		}
	};
}
