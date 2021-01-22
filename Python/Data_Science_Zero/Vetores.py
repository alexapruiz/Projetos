def vector_add(v, w):
    """soma elementos correspondentes"""
    return [v_i + w_i
    for v_i, w_i in zip(v, w)]

def vector_subtract(v, w):
    """subtrai elementos correspondentes"""
    return [v_i - w_i
    for v_i, w_i in zip(v, w)]


def vector_sum(vectors):
    return reduce(vector_add, vectors)


def scalar_multiply(c, v):
    """c é um número, v é um vetor"""
    return [c * v_i for v_i in v]


def vector_mean(vectors):
    """computar o vetor cujo i-ésimo elemento seja a média dos
    i-ésimos elementos dos vetores inclusos"""
    n = len(vectors)
    return scalar_multiply(1/n, vector_sum(vectors))


def dot(v, w):
    """v_1 * w_1 + ... + v_n * w_n"""
    return sum(v_i * w_i for v_i, w_i in zip(v, w))


def sum_of_squares(v):
    """v_1 * v_1 + ... + v_n * v_n"""
    return dot(v, v)


import math
def magnitude(v):
    return math.sqrt(sum_of_squares(v)) # math.sqrt é a função da raiz quadrada


def squared_distance(v, w):
    """(v_1 - w_1) ** 2 + ... + (v_n - w_n) ** 2"""
    return sum_of_squares(vector_subtract(v, w))


def distance(v, w):
    return magnitude(vector_subtract(v, w))
