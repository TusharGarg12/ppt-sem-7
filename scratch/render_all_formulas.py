import os
import matplotlib
import matplotlib.pyplot as plt

os.makedirs('scratch/formulas', exist_ok=True)

# Configure matplotlib mathtext to use Computer Modern (authentic LaTeX font)
matplotlib.rcParams['mathtext.fontset'] = 'cm'
matplotlib.rcParams['font.family'] = 'STIXGeneral'

def render_math_to_file(lines, filename, figsize=(7.5, 0.7), dpi=300):
    fig = plt.figure(figsize=figsize, dpi=dpi)
    fig.patch.set_alpha(0.0)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis('off')
    ax.patch.set_alpha(0.0)
    
    n_lines = len(lines)
    if n_lines == 1:
        text, fsize, color = lines[0]
        ax.text(0.5, 0.5, text, fontsize=fsize, ha='center', va='center', color=color)
    else:
        y_positions = [0.72, 0.28] if n_lines == 2 else [0.8, 0.5, 0.2]
        for (text, fsize, color), y_pos in zip(lines, y_positions):
            ax.text(0.5, y_pos, text, fontsize=fsize, ha='center', va='center', color=color)
            
    out_path = os.path.join('scratch/formulas', filename)
    plt.savefig(out_path, dpi=dpi, transparent=True, bbox_inches='tight', pad_inches=0.01)
    plt.close()
    print(f'Successfully generated {filename}')

# ==============================================================================
# 1. Slide 5: Research Foundations
# ==============================================================================
# Box 1: Fast Adapter
render_math_to_file(
    [(r'$\Delta \theta_t = \gamma \, \Delta \theta_{t-1} + (1 - \gamma) \, \nabla_\theta \mathcal{L}_t$', 18, '#0F172A')],
    'slide5_fast_adapter.png',
    figsize=(6.2, 0.6)
)

# Box 2: Slow Memory
render_math_to_file(
    [(r'$\hat{y}_t = f(x_t; \, \theta_{\mathrm{slow}}, \, \theta_{\mathrm{fast}}), \quad \mathcal{M}_t(k) = \mathrm{CosineSim}(k, \, q_t)$', 17, '#0F172A')],
    'slide5_slow_memory.png',
    figsize=(6.8, 0.6)
)

# Box 3: Quantization & Loss
render_math_to_file(
    [(r'$\tilde{x}_t = \mathrm{round}\left( \frac{x_t - \mu}{\sigma} \cdot B \right), \quad \mathcal{L}_{\mathrm{CE}} = - \sum_{k} y_{t,k} \log \hat{p}_{t,k}$', 17, '#0F172A')],
    'slide5_quantization.png',
    figsize=(6.8, 0.65)
)

# Box 4: KernelSynth GP Composition
render_math_to_file(
    [(r'$k(t, t^\prime) = k_{\mathrm{RBF}} + k_{\mathrm{Periodic}} + k_{\mathrm{Linear}}, \quad x_{\mathrm{synth}} \sim \mathcal{GP}(0, \, k)$', 17, '#0F172A')],
    'slide5_kernelsynth.png',
    figsize=(6.8, 0.6)
)

# ==============================================================================
# 2. Slide 10: Synthetic Data Generation
# ==============================================================================
# Festival Decay Kernel (2-line formula)
render_math_to_file(
    [
        (r'$E_t = \sum_k a_k \cdot \exp\left( - \frac{|t - \tau_k|}{b_k} \right)$', 17, '#0F172A'),
        (r'$a_{\mathrm{Diwali}} = 0.70 \; (-30\%), \quad a_{\mathrm{Holi}} = 0.60 \; (-40\%), \quad b_k = 2.5 \text{ days}$', 14, '#334155')
    ],
    'slide10_festival_combined.png',
    figsize=(6.5, 0.95)
)

# Stochastic Distribution & Mean (Single Line for Maximum Horizontal Scale!)
render_math_to_file(
    [(r'$\mathrm{Headcount}_{t,h,m} \sim \mathrm{NegBin}(\mu_{t,h,m}, \, \alpha_t), \quad \mu = \mathrm{Cap}_h \cdot \mathrm{Base}_m \cdot \prod_j M_j(t)$', 16, '#0F172A')],
    'slide10_stochastic.png',
    figsize=(7.2, 0.6)
)

# Heteroscedastic Dispersion
render_math_to_file(
    [(r'$\mathrm{Var}(Y) = \mu + \alpha_t \mu^2, \quad \alpha_t = 0.05 + 0.12 \cdot \sigma(\mathrm{is\_exam\_or\_fest})$', 17, '#0F172A')],
    'slide10_dispersion.png',
    figsize=(6.8, 0.6)
)

# ==============================================================================
# 3. Slide 16: Empirical ML Implementation
# ==============================================================================
# Pinball Loss
render_math_to_file(
    [(r'$\mathcal{L}_q(y, \hat{y}) = \max\left( q(y - \hat{y}), \; (1 - q)(\hat{y} - y) \right), \quad q \in \{0.10, \, 0.50, \, 0.90\}$', 17, '#0F172A')],
    'slide16_pinball.png',
    figsize=(6.8, 0.6)
)

# ==============================================================================
# 4. Slide 20: Empirical Findings: Tolerances & Metrics
# ==============================================================================
# WAPE vs. MAPE (Compact elegant summations for maximum prominence)
render_math_to_file(
    [(r'$\mathrm{WAPE} = \frac{\sum |y - \hat{y}|}{\sum y} = 25.95\%, \qquad \mathrm{MAPE} = \frac{1}{n} \sum \frac{|y - \hat{y}|}{y} = 32.14\%$', 17, '#0F172A')],
    'slide20_wape_mape.png',
    figsize=(7.2, 0.6)
)

print('ALL OPTIMIZED FORMULAS GENERATED!')
