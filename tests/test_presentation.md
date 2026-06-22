# Embodied & Spatial AI
Spatial Intelligence and Autonomous Action
Advanced Robotics & World Modeling

---

## The Core Paradigm
"Embodied AI bridges the gap between passive digital reasoning and physical action. By grounding intelligence in 3D space, agents learn to perceive, simulate, and navigate the physical world."

---

## 4 Core Domains of Spatial AI
- **Embodied AI**: Bringing digital intelligence into physical robots to act, learn, and manipulate objects.
- **Spatial Perception**: Voxel-level 3D semantic mapping, real-time depth tracing, and active sensor fusion.
- **Autonomous Navigation**: Reactive trajectory optimization, localized SLAM, and dynamic obstacle avoidance.
- **Real-World Impact**: Deploying agents to automated assembly lines, warehouses, and complex physical environments.

---

## Embodied Agent Evolution
1. **Simulation Bootstrapping**: Warm-start agent policy networks in photorealistic synthetic virtual environments.
2. **Sim-to-Real Transfer**: Bridge the virtual-physical domain gaps using zero-shot domain randomization techniques.
3. **Physical Autonomous Mastery**: Deploy agents to refine behaviors via self-supervised online reinforcement learning.

---

## World Action Model (WAM)
- **Spatiotemporal Cross-Attention**: Processes history spatial tokens and predicts subsequent action states.
- **Spatial Register Tokens**: Encodes 3D physical coordinates into unified embedding representations.
- **Physics Constraint Bounds**: Enforces structural collision limits and gravitational constants directly in latent space.

![World Action Model Architecture](assets/model_architecture.jpg)

<!-- notes: Emphasize that WAM runs at 60fps local inference to enable real-time reactive correction. -->