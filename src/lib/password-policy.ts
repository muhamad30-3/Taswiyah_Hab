const minimumPasswordLength = 6
const maximumPasswordLength = 256

export function passwordPolicyError(password: string): string | null {
  const length = Array.from(password).length
  if (length < minimumPasswordLength || length > maximumPasswordLength) {
    return "كلمة المرور يجب أن تتكون من 6 إلى 256 محرفاً."
  }
  return null
}
