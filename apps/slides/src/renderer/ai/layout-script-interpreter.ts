import { parse } from 'acorn'

import { compileBoundedRegex, type BoundedRegex } from './bounded-regex'

type AstNode = {
  type: string
  start?: number
  loc?: { start: { line: number; column: number } }
  [key: string]: unknown
}

type ScriptValue = unknown
type HostCall = (...args: ScriptValue[]) => ScriptValue

const MAX_STEPS = 25_000
const MAX_CALL_DEPTH = 64
const MAX_COLLECTION_SIZE = 10_000

class Builtin {
  constructor(
    readonly call: HostCall,
    readonly members?: Record<string, ScriptValue>,
  ) {}
}

class ScriptFunction {
  constructor(
    readonly params: AstNode[],
    readonly body: AstNode,
    readonly closure: Scope,
    readonly expressionBody: boolean,
  ) {}
}

class RegexValue {
  readonly matcher: BoundedRegex

  constructor(source: string, flags: string) {
    this.matcher = compileBoundedRegex(source, flags)
  }
}

class Scope {
  private readonly values = new Map<string, ScriptValue>()

  constructor(private readonly parent?: Scope) {}

  declare(name: string, value: ScriptValue): void {
    if (this.values.has(name)) throw new Error(`Identifier "${name}" has already been declared`)
    this.values.set(name, value)
  }

  get(name: string): ScriptValue {
    if (this.values.has(name)) return this.values.get(name)
    if (this.parent) return this.parent.get(name)
    throw new Error(
      `Unknown identifier "${name}". Only the documented layout-script API is available.`,
    )
  }

  has(name: string): boolean {
    return this.values.has(name) || Boolean(this.parent?.has(name))
  }

  set(name: string, value: ScriptValue): void {
    if (this.values.has(name)) {
      this.values.set(name, value)
      return
    }
    if (this.parent) {
      this.parent.set(name, value)
      return
    }
    throw new Error(`Cannot assign to unknown identifier "${name}"`)
  }
}

class ReturnSignal {
  constructor(readonly value: ScriptValue) {}
}
class BreakSignal {}
class ContinueSignal {}

function isNode(value: unknown): value is AstNode {
  return !!value && typeof value === 'object' && typeof (value as AstNode).type === 'string'
}

function propertyKey(value: ScriptValue): string {
  if (typeof value !== 'string' && typeof value !== 'number') {
    throw new Error('Computed property names must resolve to a string or number')
  }
  return String(value)
}

function arrayIndex(key: string): number | null {
  if (!/^(0|[1-9]\d*)$/.test(key)) return null
  const value = Number(key)
  return Number.isSafeInteger(value) ? value : null
}

function ownRecord(value: ScriptValue): value is Record<string, ScriptValue> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  return Object.getPrototypeOf(value) === null
}

/** Convert host input/output to prototype-free JSON-like data before scripts can inspect it. */
function cloneData(value: ScriptValue, depth = 0): ScriptValue {
  if (depth > MAX_CALL_DEPTH) throw new Error('Layout-script data is nested too deeply')
  if (
    value === undefined ||
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  )
    return value
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('Layout-script data contains a non-finite number')
    return value
  }
  if (Array.isArray(value)) {
    if (value.length > MAX_COLLECTION_SIZE) throw new Error('Layout-script array is too large')
    return value.map((item) => cloneData(item, depth + 1))
  }
  if (typeof value === 'object') {
    const out: Record<string, ScriptValue> = Object.create(null)
    for (const [key, item] of Object.entries(value)) out[key] = cloneData(item, depth + 1)
    return out
  }
  throw new Error(`Unsupported layout-script value type: ${typeof value}`)
}

function displayValue(value: ScriptValue): string {
  if (typeof value === 'string') return value
  try {
    return JSON.stringify(value)
  } catch {
    return String(value)
  }
}

export interface LayoutInterpreterGlobals {
  els: unknown[]
  canvas: { w: number; h: number }
  setBox: HostCall
  moveBy: HostCall
  resizeBy: HostCall
  setText: HostCall
  setStyle: HostCall
  setFill: HostCall
  setStroke: HostCall
  log: HostCall
}

/**
 * Parse and execute the layout DSL without invoking JavaScript. The evaluator implements a
 * deliberately small JavaScript-shaped language; every property read and call is dispatched by an
 * explicit allowlist, so host prototypes, constructors, globals, modules, and network APIs are not
 * representable in the execution model.
 */
export function interpretLayoutScript(
  code: string,
  globals: LayoutInterpreterGlobals,
): ScriptValue {
  let program: AstNode
  try {
    program = parse(code, {
      ecmaVersion: 'latest',
      sourceType: 'script',
      allowReturnOutsideFunction: true,
      locations: true,
    }) as unknown as AstNode
  } catch (error) {
    throw new Error(error instanceof Error ? error.message : String(error), { cause: error })
  }

  let steps = 0
  let callDepth = 0
  const tick = (node: AstNode): void => {
    steps += 1
    if (steps > MAX_STEPS) {
      const line = node.loc?.start.line
      throw new Error(
        `Layout script exceeded the ${MAX_STEPS.toLocaleString()} step limit${line ? ` near line ${line}` : ''}`,
      )
    }
  }

  const root = new Scope()
  const exposeCall = (name: keyof LayoutInterpreterGlobals): void => {
    const call = globals[name] as HostCall
    root.declare(name, new Builtin((...args) => call(...args.map((arg) => cloneData(arg)))))
  }
  root.declare('els', cloneData(globals.els))
  root.declare('canvas', cloneData(globals.canvas))
  for (const name of [
    'setBox',
    'moveBy',
    'resizeBy',
    'setText',
    'setStyle',
    'setFill',
    'setStroke',
    'log',
  ] as const) {
    exposeCall(name)
  }

  const math: Record<string, ScriptValue> = Object.create(null)
  for (const name of [
    'abs',
    'ceil',
    'floor',
    'round',
    'trunc',
    'max',
    'min',
    'pow',
    'sqrt',
    'sign',
  ] as const) {
    const call = Math[name] as (...values: number[]) => number
    math[name] = new Builtin((...args) => call(...args.map(Number)))
  }
  Object.assign(math, { PI: Math.PI, E: Math.E })
  root.declare('Math', math)

  const number: Record<string, ScriptValue> = Object.create(null)
  number.isFinite = new Builtin((value) => typeof value === 'number' && Number.isFinite(value))
  number.isInteger = new Builtin((value) => typeof value === 'number' && Number.isInteger(value))
  root.declare('Number', new Builtin((value) => Number(value), number))
  root.declare('String', new Builtin((value) => String(value)))
  root.declare('Boolean', new Builtin((value) => Boolean(value)))

  const object: Record<string, ScriptValue> = Object.create(null)
  object.keys = new Builtin((value) => (ownRecord(value) ? Object.keys(value) : []))
  object.values = new Builtin((value) => (ownRecord(value) ? Object.values(value) : []))
  object.entries = new Builtin((value) =>
    ownRecord(value) ? Object.entries(value).map(([k, v]) => [k, v]) : [],
  )
  root.declare('Object', object)

  const array: Record<string, ScriptValue> = Object.create(null)
  array.isArray = new Builtin((value) => Array.isArray(value))
  root.declare('Array', array)

  const json: Record<string, ScriptValue> = Object.create(null)
  json.stringify = new Builtin((value) => JSON.stringify(value))
  json.parse = new Builtin((value) => cloneData(JSON.parse(String(value))))
  root.declare('JSON', json)

  const propertyName = (node: AstNode, scope: Scope): string => {
    if (node.computed) return propertyKey(evaluate(node.property as AstNode, scope))
    const property = node.property as AstNode
    if (property.type !== 'Identifier') throw new Error('Unsupported property syntax')
    return String(property.name)
  }

  const invoke = (callable: ScriptValue, args: ScriptValue[]): ScriptValue => {
    tick(program)
    if (callable instanceof Builtin) return callable.call(...args)
    if (!(callable instanceof ScriptFunction))
      throw new Error('Only documented functions and safe collection methods can be called')
    callDepth += 1
    if (callDepth > MAX_CALL_DEPTH)
      throw new Error(`Layout script exceeded the call-depth limit (${MAX_CALL_DEPTH})`)
    try {
      const scope = new Scope(callable.closure)
      callable.params.forEach((param, index) => {
        if (param.type !== 'Identifier')
          throw new Error('Function parameters must be simple identifiers')
        scope.declare(String(param.name), args[index])
      })
      if (callable.expressionBody) return evaluate(callable.body, scope)
      try {
        execute(callable.body, scope)
      } catch (signal) {
        if (signal instanceof ReturnSignal) return signal.value
        throw signal
      }
      return undefined
    } finally {
      callDepth -= 1
    }
  }

  const arrayMethod = (target: ScriptValue[], key: string): Builtin | undefined => {
    const callback = (fn: ScriptValue, value: ScriptValue, index: number): ScriptValue =>
      invoke(fn, [value, index, target])
    switch (key) {
      case 'forEach':
        return new Builtin((fn) => {
          target.forEach((v, i) => callback(fn, v, i))
          return undefined
        })
      case 'map':
        return new Builtin((fn) => target.map((v, i) => callback(fn, v, i)))
      case 'filter':
        return new Builtin((fn) => target.filter((v, i) => Boolean(callback(fn, v, i))))
      case 'find':
        return new Builtin((fn) => target.find((v, i) => Boolean(callback(fn, v, i))))
      case 'findIndex':
        return new Builtin((fn) => target.findIndex((v, i) => Boolean(callback(fn, v, i))))
      case 'some':
        return new Builtin((fn) => target.some((v, i) => Boolean(callback(fn, v, i))))
      case 'every':
        return new Builtin((fn) => target.every((v, i) => Boolean(callback(fn, v, i))))
      case 'reduce':
        return new Builtin((...args) => {
          const [fn, initial] = args
          if (args.length > 1)
            return target.reduce((acc, v, i) => invoke(fn, [acc, v, i, target]), initial)
          if (target.length === 0) throw new Error('Reduce of empty array with no initial value')
          return target
            .slice(1)
            .reduce((acc, v, i) => invoke(fn, [acc, v, i + 1, target]), target[0])
        })
      case 'slice':
        return new Builtin((start, end) =>
          target.slice(
            start === undefined ? undefined : Number(start),
            end === undefined ? undefined : Number(end),
          ),
        )
      case 'indexOf':
        return new Builtin((value) => target.indexOf(value))
      case 'includes':
        return new Builtin((value) => target.includes(value))
      case 'join':
        return new Builtin((separator) =>
          target.map(displayValue).join(separator === undefined ? ',' : String(separator)),
        )
      case 'push':
        return new Builtin((...values) => {
          if (target.length + values.length > MAX_COLLECTION_SIZE)
            throw new Error('Layout-script array is too large')
          return target.push(...values)
        })
      case 'pop':
        return new Builtin(() => target.pop())
      case 'shift':
        return new Builtin(() => target.shift())
      case 'unshift':
        return new Builtin((...values) => target.unshift(...values))
      case 'reverse':
        return new Builtin(() => target.reverse())
      case 'sort':
        return new Builtin((fn) =>
          target.sort(fn === undefined ? undefined : (a, b) => Number(invoke(fn, [a, b]))),
        )
      default:
        return undefined
    }
  }

  const stringMethod = (target: string, key: string): Builtin | undefined => {
    switch (key) {
      case 'includes':
        return new Builtin((value) => target.includes(String(value)))
      case 'startsWith':
        return new Builtin((value) => target.startsWith(String(value)))
      case 'endsWith':
        return new Builtin((value) => target.endsWith(String(value)))
      case 'toLowerCase':
        return new Builtin(() => target.toLowerCase())
      case 'toUpperCase':
        return new Builtin(() => target.toUpperCase())
      case 'trim':
        return new Builtin(() => target.trim())
      case 'slice':
        return new Builtin((start, end) =>
          target.slice(Number(start ?? 0), end === undefined ? undefined : Number(end)),
        )
      case 'substring':
        return new Builtin((start, end) =>
          target.substring(Number(start ?? 0), end === undefined ? undefined : Number(end)),
        )
      case 'indexOf':
        return new Builtin((value) => target.indexOf(String(value)))
      case 'split':
        return new Builtin((separator, limit) =>
          separator === undefined
            ? [target]
            : target.split(String(separator), limit === undefined ? undefined : Number(limit)),
        )
      default:
        return undefined
    }
  }

  const getMember = (target: ScriptValue, key: string): ScriptValue => {
    if (Array.isArray(target)) {
      if (key === 'length') return target.length
      const index = arrayIndex(key)
      if (index !== null) return target[index]
      const method = arrayMethod(target, key)
      if (method) return method
      throw new Error(`Array property "${key}" is not available in layout scripts`)
    }
    if (typeof target === 'string') {
      if (key === 'length') return target.length
      const index = arrayIndex(key)
      if (index !== null) return target[index]
      const method = stringMethod(target, key)
      if (method) return method
      throw new Error(`String property "${key}" is not available in layout scripts`)
    }
    if (target instanceof RegexValue) {
      if (key === 'test') return new Builtin((value) => target.matcher.test(String(value)))
      throw new Error(`Regular-expression property "${key}" is not available in layout scripts`)
    }
    if (target instanceof Builtin) {
      if (target.members && Object.prototype.hasOwnProperty.call(target.members, key))
        return target.members[key]
      throw new Error(`Function property "${key}" is not available in layout scripts`)
    }
    if (ownRecord(target)) {
      if (Object.prototype.hasOwnProperty.call(target, key)) return target[key]
      return undefined
    }
    if (target === null || target === undefined)
      throw new Error(`Cannot read property "${key}" of ${target}`)
    throw new Error(`Properties are not available on ${typeof target} values`)
  }

  const setMember = (target: ScriptValue, key: string, value: ScriptValue): ScriptValue => {
    if (Array.isArray(target)) {
      const index = arrayIndex(key)
      if (index === null || index > MAX_COLLECTION_SIZE)
        throw new Error(`Array property "${key}" cannot be assigned`)
      target[index] = value
      return value
    }
    if (ownRecord(target)) {
      target[key] = value
      return value
    }
    throw new Error(`Property "${key}" cannot be assigned`)
  }

  const binary = (operator: string, left: ScriptValue, right: ScriptValue): ScriptValue => {
    switch (operator) {
      case '+':
        return (left as never) + (right as never)
      case '-':
        return Number(left) - Number(right)
      case '*':
        return Number(left) * Number(right)
      case '/':
        return Number(left) / Number(right)
      case '%':
        return Number(left) % Number(right)
      case '**':
        return Number(left) ** Number(right)
      case '<':
        return (left as never) < (right as never)
      case '<=':
        return (left as never) <= (right as never)
      case '>':
        return (left as never) > (right as never)
      case '>=':
        return (left as never) >= (right as never)
      case '==':
        return left == right // DSL compatibility with ordinary layout snippets
      case '!=':
        return left != right
      case '===':
        return left === right
      case '!==':
        return left !== right
      default:
        throw new Error(`Binary operator "${operator}" is not supported`)
    }
  }

  const assign = (node: AstNode, scope: Scope, value: ScriptValue): ScriptValue => {
    if (node.type === 'Identifier') {
      scope.set(String(node.name), value)
      return value
    }
    if (node.type === 'MemberExpression') {
      const target = evaluate(node.object as AstNode, scope)
      return setMember(target, propertyName(node, scope), value)
    }
    throw new Error('Invalid assignment target')
  }

  const evaluate = (node: AstNode, scope: Scope): ScriptValue => {
    tick(node)
    switch (node.type) {
      case 'Literal': {
        const regex = node.regex as { pattern: string; flags: string } | undefined
        return regex ? new RegexValue(regex.pattern, regex.flags) : node.value
      }
      case 'Identifier':
        return scope.get(String(node.name))
      case 'ArrayExpression': {
        const values: ScriptValue[] = []
        for (const item of node.elements as Array<AstNode | null>) {
          if (!item) {
            values.push(undefined)
            continue
          }
          if (item.type === 'SpreadElement') {
            const spread = evaluate(item.argument as AstNode, scope)
            if (!Array.isArray(spread)) throw new Error('Array spread requires an array')
            values.push(...spread)
          } else values.push(evaluate(item, scope))
        }
        return values
      }
      case 'ObjectExpression': {
        const out: Record<string, ScriptValue> = Object.create(null)
        for (const property of node.properties as AstNode[]) {
          if (property.type === 'SpreadElement') {
            const spread = evaluate(property.argument as AstNode, scope)
            if (!ownRecord(spread)) throw new Error('Object spread requires a plain data object')
            Object.assign(out, spread)
            continue
          }
          if (property.type !== 'Property' || property.kind !== 'init' || property.method) {
            throw new Error('Only data properties are supported in object literals')
          }
          const key = property.computed
            ? propertyKey(evaluate(property.key as AstNode, scope))
            : String((property.key as AstNode).name ?? (property.key as AstNode).value)
          out[key] = evaluate(property.value as AstNode, scope)
        }
        return out
      }
      case 'UnaryExpression': {
        const argument = node.argument as AstNode
        if (
          node.operator === 'typeof' &&
          argument.type === 'Identifier' &&
          !scope.has(String(argument.name))
        ) {
          return 'undefined'
        }
        const value = evaluate(argument, scope)
        switch (node.operator) {
          case '!':
            return !value
          case '+':
            return Number(value)
          case '-':
            return -Number(value)
          case 'typeof':
            return typeof value
          default:
            throw new Error(`Unary operator "${String(node.operator)}" is not supported`)
        }
      }
      case 'BinaryExpression':
        return binary(
          String(node.operator),
          evaluate(node.left as AstNode, scope),
          evaluate(node.right as AstNode, scope),
        )
      case 'LogicalExpression': {
        const left = evaluate(node.left as AstNode, scope)
        if (node.operator === '&&') return left ? evaluate(node.right as AstNode, scope) : left
        if (node.operator === '||') return left ? left : evaluate(node.right as AstNode, scope)
        if (node.operator === '??')
          return left === null || left === undefined ? evaluate(node.right as AstNode, scope) : left
        throw new Error(`Logical operator "${String(node.operator)}" is not supported`)
      }
      case 'ConditionalExpression':
        return evaluate(
          (evaluate(node.test as AstNode, scope) ? node.consequent : node.alternate) as AstNode,
          scope,
        )
      case 'MemberExpression': {
        const target = evaluate(node.object as AstNode, scope)
        if (node.optional && (target === null || target === undefined)) return undefined
        return getMember(target, propertyName(node, scope))
      }
      case 'ChainExpression':
        return evaluate(node.expression as AstNode, scope)
      case 'CallExpression': {
        const callable = evaluate(node.callee as AstNode, scope)
        if (node.optional && (callable === null || callable === undefined)) return undefined
        const args: ScriptValue[] = []
        for (const argument of node.arguments as AstNode[]) {
          if (argument.type === 'SpreadElement') {
            const spread = evaluate(argument.argument as AstNode, scope)
            if (!Array.isArray(spread)) throw new Error('Call spread requires an array')
            args.push(...spread)
          } else args.push(evaluate(argument, scope))
        }
        return invoke(callable, args)
      }
      case 'ArrowFunctionExpression':
      case 'FunctionExpression':
        return new ScriptFunction(
          node.params as AstNode[],
          node.body as AstNode,
          scope,
          Boolean(node.body && (node.body as AstNode).type !== 'BlockStatement'),
        )
      case 'AssignmentExpression': {
        const left = node.left as AstNode
        if (node.operator === '=')
          return assign(left, scope, evaluate(node.right as AstNode, scope))
        const current = evaluate(left, scope)
        const operator = String(node.operator).slice(0, -1)
        return assign(
          left,
          scope,
          binary(operator, current, evaluate(node.right as AstNode, scope)),
        )
      }
      case 'UpdateExpression': {
        const argument = node.argument as AstNode
        const current = Number(evaluate(argument, scope))
        const next = node.operator === '++' ? current + 1 : current - 1
        assign(argument, scope, next)
        return node.prefix ? next : current
      }
      case 'TemplateLiteral': {
        const expressions = node.expressions as AstNode[]
        return (node.quasis as AstNode[])
          .map((quasi, index) => {
            const cooked = (quasi.value as { cooked?: string }).cooked ?? ''
            return (
              cooked +
              (index < expressions.length ? String(evaluate(expressions[index], scope)) : '')
            )
          })
          .join('')
      }
      case 'SequenceExpression': {
        let value: ScriptValue
        for (const expression of node.expressions as AstNode[]) value = evaluate(expression, scope)
        return value
      }
      default:
        throw new Error(`Expression type "${node.type}" is not supported in layout scripts`)
    }
  }

  const execute = (node: AstNode, scope: Scope): void => {
    tick(node)
    switch (node.type) {
      case 'Program':
      case 'BlockStatement': {
        const blockScope = node.type === 'Program' ? scope : new Scope(scope)
        for (const statement of node.body as AstNode[]) execute(statement, blockScope)
        return
      }
      case 'EmptyStatement':
        return
      case 'ExpressionStatement':
        evaluate(node.expression as AstNode, scope)
        return
      case 'VariableDeclaration':
        for (const declaration of node.declarations as AstNode[]) {
          const id = declaration.id as AstNode
          const value = declaration.init ? evaluate(declaration.init as AstNode, scope) : undefined
          if (id.type === 'Identifier') {
            scope.declare(String(id.name), value)
            continue
          }
          if (id.type === 'ArrayPattern') {
            if (!Array.isArray(value)) throw new Error('Array destructuring requires an array')
            for (const [index, item] of (id.elements as Array<AstNode | null>).entries()) {
              if (!item) continue
              if (item.type !== 'Identifier')
                throw new Error('Nested or rest destructuring is not supported')
              scope.declare(String(item.name), value[index])
            }
            continue
          }
          throw new Error('Variable declarations must use identifiers or flat array destructuring')
        }
        return
      case 'ReturnStatement':
        throw new ReturnSignal(
          node.argument ? evaluate(node.argument as AstNode, scope) : undefined,
        )
      case 'IfStatement':
        if (evaluate(node.test as AstNode, scope)) execute(node.consequent as AstNode, scope)
        else if (node.alternate) execute(node.alternate as AstNode, scope)
        return
      case 'ForStatement': {
        const loopScope = new Scope(scope)
        if (node.init) {
          if (isNode(node.init) && node.init.type === 'VariableDeclaration')
            execute(node.init, loopScope)
          else evaluate(node.init as AstNode, loopScope)
        }
        while (!node.test || evaluate(node.test as AstNode, loopScope)) {
          try {
            execute(node.body as AstNode, loopScope)
          } catch (signal) {
            if (signal instanceof BreakSignal) break
            if (!(signal instanceof ContinueSignal)) throw signal
          }
          if (node.update) evaluate(node.update as AstNode, loopScope)
        }
        return
      }
      case 'ForOfStatement': {
        const values = evaluate(node.right as AstNode, scope)
        if (!Array.isArray(values) && typeof values !== 'string')
          throw new Error('for...of requires an array or string')
        for (const value of values) {
          const loopScope = new Scope(scope)
          const left = node.left as AstNode
          if (left.type === 'VariableDeclaration') {
            const declaration = (left.declarations as AstNode[])[0]
            const id = declaration.id as AstNode
            if (id.type !== 'Identifier')
              throw new Error('for...of variables must be simple identifiers')
            loopScope.declare(String(id.name), value)
          } else assign(left, loopScope, value)
          try {
            execute(node.body as AstNode, loopScope)
          } catch (signal) {
            if (signal instanceof BreakSignal) break
            if (!(signal instanceof ContinueSignal)) throw signal
          }
        }
        return
      }
      case 'WhileStatement':
        while (evaluate(node.test as AstNode, scope)) {
          try {
            execute(node.body as AstNode, scope)
          } catch (signal) {
            if (signal instanceof BreakSignal) break
            if (!(signal instanceof ContinueSignal)) throw signal
          }
        }
        return
      case 'BreakStatement':
        throw new BreakSignal()
      case 'ContinueStatement':
        throw new ContinueSignal()
      case 'FunctionDeclaration': {
        const id = node.id as AstNode | null
        if (!id || id.type !== 'Identifier') throw new Error('Functions must have a simple name')
        scope.declare(
          String(id.name),
          new ScriptFunction(node.params as AstNode[], node.body as AstNode, scope, false),
        )
        return
      }
      case 'ThrowStatement':
        throw new Error(displayValue(evaluate(node.argument as AstNode, scope)))
      default:
        throw new Error(`Statement type "${node.type}" is not supported in layout scripts`)
    }
  }

  try {
    execute(program, root)
  } catch (signal) {
    if (signal instanceof ReturnSignal) return cloneData(signal.value)
    throw signal
  }
  return undefined
}
