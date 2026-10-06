const minimumPasswordLength = 8
const maximumPasswordLength = 256

export function passwordPolicyError(password: string): string | null {
  const length = Array.from(password).length
  if (length < minimumPasswordLength || length > maximumPasswordLength) {
    return "كلمة المرور يجب أن تتكون من 8 إلى 256 محرفاً."
  }
  if (!/\p{L}/u.test(password) || !/\p{N}/u.test(password)) {
    return "يجب أن تحتوي كلمة المرور على حرف واحد ورقم واحد على الأقل."
  }
  return null
}
