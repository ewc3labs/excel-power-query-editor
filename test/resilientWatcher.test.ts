import * as assert from 'assert';
import { EventEmitter } from 'events';
import { createResilientWatcher, ResilientWatcherOptions } from '../src/resilientWatcher';

/**
 * The fallback STATE MACHINE, with an injected watcher.
 *
 * This is not a fake of the network share - whether polling works on somebody's file server is a
 * question only a real share answers (PQ-36, still coded). What these tests own is the part that was
 * shipped untested the first time: swap exactly once, hand the holder the live watcher, ignore
 * errors from a watcher that has been replaced or closed, and stop - not loop - when polling fails.
 */

class FakeWatcher extends EventEmitter {
	closed = false;
	constructor(readonly polling: boolean) { super(); }
	async close(): Promise<void> { this.closed = true; }
	fail(message = 'UNKNOWN: unknown error, watch'): void { this.emit('error', new Error(message)); }
}

const tick = () => new Promise<void>((resolve) => setImmediate(resolve));

function harness(startPolling = false) {
	const created: FakeWatcher[] = [];
	const attached: { polling: boolean }[] = [];
	const replaced: FakeWatcher[] = [];
	const unwatchable: unknown[] = [];
	const logs: string[] = [];

	const opts: ResilientWatcherOptions<FakeWatcher> = {
		startPolling,
		create: (usePolling) => { const w = new FakeWatcher(usePolling); created.push(w); return w; },
		attach: (_w, usePolling) => { attached.push({ polling: usePolling }); },
		onReplaced: (w) => { replaced.push(w); },
		onUnwatchable: (e) => { unwatchable.push(e); },
		log: (m) => { logs.push(m); }
	};
	const watcher = createResilientWatcher(opts);
	return { watcher, created, attached, replaced, unwatchable, logs };
}

suite('Resilient file watcher (PQ-36)', () => {
	test('starts native, with handlers attached', () => {
		const h = harness();
		assert.strictEqual(h.created.length, 1);
		assert.strictEqual(h.created[0].polling, false);
		assert.deepStrictEqual(h.attached, [{ polling: false }]);
		assert.strictEqual(h.watcher.state, 'native');
		assert.strictEqual(h.watcher.current, h.created[0]);
	});

	test('a native failure swaps to polling once, repoints the holder, and closes the dead watcher', async () => {
		const h = harness();
		const native = h.created[0];

		native.fail();
		await tick();

		assert.strictEqual(h.created.length, 2, 'exactly one replacement');
		const polling = h.created[1];
		assert.strictEqual(polling.polling, true);
		assert.deepStrictEqual(h.attached, [{ polling: false }, { polling: true }], 'the replacement gets the same handlers');
		assert.deepStrictEqual(h.replaced, [polling], 'whoever holds the old watcher must hear about the new one');
		assert.strictEqual(h.watcher.current, polling);
		assert.strictEqual(h.watcher.state, 'polling');
		assert.strictEqual(native.closed, true, 'the dead watcher is closed, not leaked');
		assert.strictEqual(polling.closed, false);
		assert.strictEqual(h.unwatchable.length, 0);
	});

	test('two native errors in the same tick start only one fallback', async () => {
		// A failing native watcher rarely fails exactly once. The transition is claimed before any
		// await, so the second error must not build a second polling watcher.
		const h = harness();
		h.created[0].fail();
		h.created[0].fail();
		await tick();

		assert.strictEqual(h.created.length, 2);
		assert.strictEqual(h.replaced.length, 1);
	});

	test('a late error from the replaced watcher is ignored', async () => {
		const h = harness();
		const native = h.created[0];
		native.fail();
		await tick();

		native.fail('late');
		await tick();

		assert.strictEqual(h.created.length, 2, 'no third watcher');
		assert.strictEqual(h.watcher.state, 'polling', 'the working polling watcher is not stopped by the dead one');
		assert.strictEqual(h.unwatchable.length, 0);
	});

	test('a polling failure stops and reports unwatchable, instead of looping', async () => {
		const h = harness();
		h.created[0].fail();
		await tick();
		const polling = h.created[1];

		polling.fail('polling broke too');
		await tick();

		assert.strictEqual(h.created.length, 2, 'nothing left to fall back to, so nothing new is built');
		assert.strictEqual(h.watcher.state, 'stopped');
		assert.strictEqual(polling.closed, true);
		assert.strictEqual(h.unwatchable.length, 1, 'reported once');
	});

	test('starting in polling mode has nothing to fall back to', async () => {
		// Dev containers poll from the start. Their first failure is final.
		const h = harness(true);
		assert.strictEqual(h.created[0].polling, true);
		assert.strictEqual(h.watcher.state, 'polling');

		h.created[0].fail();
		await tick();

		assert.strictEqual(h.created.length, 1);
		assert.strictEqual(h.watcher.state, 'stopped');
		assert.strictEqual(h.unwatchable.length, 1);
	});

	test('an error after close does nothing - stopping watching means stopped', async () => {
		// The leak this prevents: stop watching, then the old watcher reports an error, and a polling
		// watcher starts that nothing holds and nothing will ever close.
		const h = harness();
		await h.watcher.close();
		assert.strictEqual(h.created[0].closed, true);

		h.created[0].fail();
		await tick();

		assert.strictEqual(h.created.length, 1, 'no fallback after the user stopped watching');
		assert.strictEqual(h.unwatchable.length, 0);
		assert.strictEqual(h.watcher.state, 'stopped');
	});

	test('close after a fallback closes the live polling watcher', async () => {
		const h = harness();
		h.created[0].fail();
		await tick();

		await h.watcher.close();
		assert.strictEqual(h.created[1].closed, true, 'teardown must reach the watcher that is actually running');
	});
});
