export function projectToSlide(x: number, y: number, z: number) {
  const scale = 2
  return { slideX: x * scale, slideY: y * scale - z * 0.5 }
}
