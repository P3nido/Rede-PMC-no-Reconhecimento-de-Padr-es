import os
import numpy as np
import matplotlib.pyplot as plt


# =============================================================================
# Leitura dos dados
# =============================================================================

def load_data(filepath):
    """Le arquivo .txt com cabecalho e retorna array numpy."""
    data = []
    with open(filepath, 'r') as f:
        lines = f.readlines()
    for line in lines[1:]:  # pula cabecalho
        values = line.strip().split()
        if values:
            data.append([float(v) for v in values])
    return np.array(data)


# =============================================================================
# Funcoes de ativacao
# =============================================================================

def sigmoid(x):
    return 1 / (1 + np.exp(-x))

def sigmoid_deriv(x):
    return x * (1 - x)


# =============================================================================
# Rede PMC (Perceptron Multi-Camadas)
# =============================================================================

class MLP:

    def __init__(self, input_size, hidden_size, output_size, learning_rate=0.1, seed=None):
        if seed is not None:
            np.random.seed(seed)

        # Pesos e bias inicializados aleatoriamente entre 0 e 1
        self.W1 = np.random.rand(input_size, hidden_size)
        self.b1 = np.random.rand(hidden_size)
        self.W2 = np.random.rand(hidden_size, output_size)
        self.b2 = np.random.rand(output_size)

        self.learning_rate = learning_rate

    def forward(self, X):
        z1 = np.dot(X, self.W1) + self.b1
        a1 = sigmoid(z1)
        z2 = np.dot(a1, self.W2) + self.b2
        a2 = sigmoid(z2)
        return a2

    def train(self, X, y, epsilon=1e-6, max_epochs=10000):
        eqm_list = []
        n_amostras = len(X)
        eqm_anterior = float('inf')

        for epoch in range(max_epochs):

            # Embaralha as amostras a cada epoca
            indices = np.arange(n_amostras)
            np.random.shuffle(indices)
            X_shuffled = X[indices]
            y_shuffled = y[indices]

            # Atualizacao amostra por amostra (SGD)
            for i in range(n_amostras):
                x_i = X_shuffled[i:i+1]
                y_i = y_shuffled[i:i+1]

                # Forward pass
                z1 = np.dot(x_i, self.W1) + self.b1
                a1 = sigmoid(z1)
                z2 = np.dot(a1, self.W2) + self.b2
                a2 = sigmoid(z2)

                # Erro por amostra
                erro = y_i - a2

                # Backward pass
                delta_output = erro * sigmoid_deriv(a2)
                error_hidden = np.dot(delta_output, self.W2.T)
                delta_hidden = error_hidden * sigmoid_deriv(a1)

                # Atualizacao dos pesos e bias
                self.W2 += self.learning_rate * np.dot(a1.T, delta_output)
                self.b2 += self.learning_rate * np.sum(delta_output, axis=0)
                self.W1 += self.learning_rate * np.dot(x_i.T, delta_hidden)
                self.b1 += self.learning_rate * np.sum(delta_hidden, axis=0)

            # EQM da epoca sobre todo o conjunto de treino
            todas_estimativas = self.forward(X)
            eqm_atual = np.mean((y - todas_estimativas) ** 2)
            eqm_list.append(eqm_atual)

            if epoch % 100 == 0:
                print(f"    Epoca {epoch:5d} - EQM = {eqm_atual:.8f}")

            # Criterio de parada: variacao do EQM entre epocas
            variacao = abs(eqm_atual - eqm_anterior)
            if variacao <= epsilon:
                break
            eqm_anterior = eqm_atual

        return eqm_list, epoch + 1

    def predict(self, X):
        return self.forward(X)


# =============================================================================
# Metricas de avaliacao
# =============================================================================

def confusion_matrix_3class(y_true, y_pred):
    """Matriz de confusao 3x3 baseada em argmax (one-hot -> classe)."""
    true_classes = np.argmax(y_true, axis=1)
    pred_classes = np.argmax(y_pred, axis=1)

    cm = np.zeros((3, 3), dtype=int)
    for t, p in zip(true_classes, pred_classes):
        cm[t][p] += 1
    return cm


def metrics_from_cm(cm):
    """Calcula acuracia, precisao e recall a partir da matriz de confusao."""
    total = np.sum(cm)
    acuracia = np.trace(cm) / total

    precisao = []
    recall = []
    for i in range(cm.shape[0]):
        tp = cm[i, i]
        fp = np.sum(cm[:, i]) - tp
        fn = np.sum(cm[i, :]) - tp
        p = tp / (tp + fp) if (tp + fp) > 0 else 0.0
        r = tp / (tp + fn) if (tp + fn) > 0 else 0.0
        precisao.append(p)
        recall.append(r)

    return acuracia, precisao, recall


# =============================================================================
# Graficos
# =============================================================================

def salvar_grafico_eqm(eqm_list, output_path, titulo):
    plt.figure(figsize=(7, 4.5))
    plt.plot(eqm_list, color='steelblue')
    plt.title(titulo)
    plt.xlabel("Epocas")
    plt.ylabel("EQM")
    plt.yscale("log")
    plt.grid(True)
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close()


def salvar_grafico_validacao(y_true, y_pred, output_path, titulo):
    """Plota classe predita vs classe desejada para cada amostra de validacao."""
    true_classes = np.argmax(y_true, axis=1) + 1
    pred_classes = np.argmax(y_pred, axis=1) + 1
    amostras = np.arange(1, len(y_true) + 1)

    plt.figure(figsize=(9, 5))
    plt.plot(amostras, true_classes, 'o-', label="Classe desejada", color='blue')
    plt.plot(amostras, pred_classes, 's--', label="Classe estimada", color='red')
    plt.title(titulo)
    plt.xlabel("Amostras de validacao")
    plt.ylabel("Classe")
    plt.yticks([1, 2, 3], ['Classe 1', 'Classe 2', 'Classe 3'])
    plt.grid(True)
    plt.legend()
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close()


def salvar_grafico_confusao(cm, output_path, titulo):
    """Plota a matriz de confusao como heatmap."""
    fig, ax = plt.subplots(figsize=(5, 4.5))
    im = ax.imshow(cm, interpolation='nearest', cmap=plt.cm.Blues)
    plt.colorbar(im, ax=ax)

    classes = ['Classe 1', 'Classe 2', 'Classe 3']
    tick_marks = np.arange(len(classes))
    ax.set_xticks(tick_marks)
    ax.set_xticklabels(classes, rotation=45)
    ax.set_yticks(tick_marks)
    ax.set_yticklabels(classes)

    thresh = cm.max() / 2.0
    for i in range(cm.shape[0]):
        for j in range(cm.shape[1]):
            ax.text(j, i, format(cm[i, j], 'd'),
                    ha="center", va="center",
                    color="white" if cm[i, j] > thresh else "black",
                    fontsize=14)

    ax.set_ylabel('Classe verdadeira')
    ax.set_xlabel('Classe predita')
    ax.set_title(titulo)
    plt.tight_layout()
    plt.savefig(output_path, dpi=300)
    plt.close()


# =============================================================================
# Main
# =============================================================================

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))

    treino_path    = os.path.join(base_dir, "PP04_dados_treinamento.txt")
    validacao_path = os.path.join(base_dir, "PP04_dados_validacao.txt")
    output_dir     = os.path.join(base_dir, "graphics")
    os.makedirs(output_dir, exist_ok=True)

    # Carrega dados dos arquivos
    dados_treino    = load_data(treino_path)
    dados_validacao = load_data(validacao_path)

    X_train = dados_treino[:, :4]       # entradas x1..x4
    y_train = dados_treino[:, 4:]       # saidas  d1, d2, d3

    X_val   = dados_validacao[:, :4]
    y_val   = dados_validacao[:, 4:]

    print(f"Amostras de treinamento : {X_train.shape[0]}")
    print(f"Amostras de validacao   : {X_val.shape[0]}")
    print(f"Entradas: {X_train.shape[1]}  |  Saidas: {y_train.shape[1]}\n")

    redes       = []
    eqms_finais = []
    epocas_list = []

    print("=" * 60)
    print("INICIANDO OS 5 TREINAMENTOS")
    print("=" * 60)

    for i in range(5):
        print(f"\n-- Treinamento T{i+1} --")

        mlp = MLP(input_size=4, hidden_size=10, output_size=3, learning_rate=0.1)
        eqm_list, n_epocas = mlp.train(X_train, y_train, epsilon=1e-6, max_epochs=10000)

        redes.append(mlp)
        eqms_finais.append(eqm_list[-1])
        epocas_list.append(n_epocas)

        print(f"  T{i+1}: EQM final = {eqm_list[-1]:.8f} | Epocas = {n_epocas}")

        # Grafico EQM x Epocas
        salvar_grafico_eqm(
            eqm_list,
            os.path.join(output_dir, f"eqm_treinamento_T{i+1}.png"),
            f"EQM - Treinamento T{i+1}"
        )

        # Validacao
        y_pred = mlp.predict(X_val)

        cm = confusion_matrix_3class(y_val, y_pred)
        acuracia, precisao, recall = metrics_from_cm(cm)

        print(f"\n  Validacao T{i+1}:")
        print(f"  Matriz de Confusao:\n{cm}")
        print(f"  Acuracia: {acuracia:.4f}")
        for cls in range(3):
            print(f"  Classe {cls+1} -> Precisao: {precisao[cls]:.4f}  |  Recall: {recall[cls]:.4f}")

        # Grafico classe predita vs desejada
        salvar_grafico_validacao(
            y_val, y_pred,
            os.path.join(output_dir, f"validacao_T{i+1}.png"),
            f"Validacao T{i+1}"
        )

        # Grafico matriz de confusao
        salvar_grafico_confusao(
            cm,
            os.path.join(output_dir, f"confusao_T{i+1}.png"),
            f"Matriz de Confusao - T{i+1}"
        )

    # ==========================================================================
    # Resumo final
    # ==========================================================================
    print("\n" + "=" * 60)
    print("RESUMO DOS 5 TREINAMENTOS")
    print("=" * 60)

    melhor_idx = int(np.argmin(eqms_finais))
    for i in range(5):
        marca = " <-- MELHOR" if i == melhor_idx else ""
        print(f"  T{i+1}: EQM = {eqms_finais[i]:.8f} | Epocas = {epocas_list[i]}{marca}")

    print(f"\nMelhor rede: T{melhor_idx + 1}")

    # Grafico comparativo de EQMs finais
    plt.figure(figsize=(8, 5))
    cores = ['steelblue'] * 5
    cores[melhor_idx] = 'darkorange'
    plt.bar([f"T{i+1}" for i in range(5)], eqms_finais, color=cores, edgecolor='black')
    plt.title("EQM Final por Treinamento")
    plt.xlabel("Treinamento")
    plt.ylabel("EQM Final")
    plt.grid(True, axis='y')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "eqm_comparativo.png"), dpi=300)
    plt.close()

    print("\nGraficos salvos em:", output_dir)


if __name__ == "__main__":
    main()
