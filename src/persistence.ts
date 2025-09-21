// src/persistence.ts
import defaultScene from "./scene-default.json";

export type Scene = unknown;

const keyFor = (slideId: string) => `scene:${slideId}`;

export async function getSlideId(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id,items/index");
    await context.sync();
    if (slides.items.length > 0 && slides.items[0].id) return slides.items[0].id;
    return "unknown";
  });
}

export function saveScene(slideId: string, scene: Scene): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      const k = keyFor(slideId);
      Office.context.document.settings.set(k, JSON.stringify(scene));
      Office.context.document.settings.saveAsync((res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error?.message || "settings.saveAsync failed"));
      });
    } catch (e) {
      reject(e as Error);
    }
  });
}

export function loadScene(slideId: string): Scene {
  const k = keyFor(slideId);
  const v = Office.context.document.settings.get(k);
  if (typeof v === "string" && v.length > 0) {
    try {
      return JSON.parse(v);
    } catch (_) {
      return defaultScene as Scene;
    }
  }
  return defaultScene as Scene;
}

export function defaultSceneObject(): Scene {
  return JSON.parse(JSON.stringify(defaultScene));
}
